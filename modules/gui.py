import sys
import os
import threading
from typing import List, Optional, Set, Dict, Tuple, Any
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit,
    QComboBox, QListWidget, QListWidgetItem, QGroupBox,
    QCheckBox, QProgressBar, QMessageBox, QSplitter, QTabWidget,
    QDoubleSpinBox, QDialog, QDialogButtonBox, QFileDialog,
    QMenu, QStyle, QSystemTrayIcon, QSpinBox, QTimeEdit
)
from PyQt6.QtCore import (
    Qt, pyqtSignal, QObject, QThread, QTimer, QLockFile, QStandardPaths,
    QTime
)
from PyQt6.QtGui import QAction, QFont, QIcon

from modules.config_manager import ConfigError, ConfigManager
from modules.wechat_sender import WeChatSender
from modules.table_processor import TableProcessor
from modules.schedule_manager import (
    ScheduleStore,
    ScheduleDispatcher,
    ScheduleTask,
    _new_task_id,
    _normalize_time,
    WEEKDAY_NAMES,
)
from modules.workstation_guard import (
    WorkstationGuard,
    is_workstation_locked,
)


APP_TITLE = "表格自动发送By春风予Lu"
INSTANCE_LOCK_NAME = "ExcelSendWx.lock"
# 与 installer.iss 的 Run 键值名一致，确保安装器勾选和托盘菜单勾选操作同一注册表项
_AUTOSTART_REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
_AUTOSTART_REG_NAME = "ExcelSendWx"


def _get_autostart_command() -> str:
    """获取用于注册表 Run 键的启动命令。

    打包后：直接用 exe 路径；
    开发模式：用 pythonw.exe + main.py。
    """
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if not os.path.exists(pythonw):
        pythonw = sys.executable
    main_py = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "main.py"))
    return f'"{pythonw}" "{main_py}"'


def is_autostart_enabled() -> bool:
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _AUTOSTART_REG_PATH, 0, winreg.KEY_READ) as key:
            winreg.QueryValueEx(key, _AUTOSTART_REG_NAME)
        return True
    except FileNotFoundError:
        return False
    except OSError:
        return False


def set_autostart(enabled: bool) -> bool:
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, _AUTOSTART_REG_PATH, 0, winreg.KEY_SET_VALUE
        ) as key:
            if enabled:
                winreg.SetValueEx(key, _AUTOSTART_REG_NAME, 0, winreg.REG_SZ, _get_autostart_command())
            else:
                try:
                    winreg.DeleteValue(key, _AUTOSTART_REG_NAME)
                except FileNotFoundError:
                    pass
        return True
    except OSError:
        return False


def get_application_icon():
    icon_path = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "love.ico")
    )
    icon = QIcon(icon_path)
    if icon.isNull():
        icon = QIcon(sys.executable)
    return icon


def acquire_instance_lock(lock_path=None):
    if not lock_path:
        lock_dir = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.TempLocation
        )
        lock_path = os.path.join(lock_dir or os.getcwd(), INSTANCE_LOCK_NAME)

    instance_lock = QLockFile(lock_path)
    if instance_lock.tryLock(100):
        return instance_lock

    if (
        instance_lock.removeStaleLockFile()
        and instance_lock.tryLock(100)
    ):
        return instance_lock
    return None



class WorkerSignals(QObject):
    error = pyqtSignal(str)
    progress = pyqtSignal(int)
    result = pyqtSignal(object)
    log = pyqtSignal(str)





class ExcelReadWorker(QThread):
    def __init__(self, file_path, sheet_name=None):
        super().__init__()
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.signals = WorkerSignals()
        self.stopped_event = threading.Event()

    def stop(self):
        self.stopped_event.set()

    def run(self):
        try:
            if self.stopped_event.is_set():
                return
                
            self.signals.log.emit("正在读取Excel文件...")
            
            result = self._read_excel(self.file_path)
            
            if self.stopped_event.is_set():
                return
                
            self.signals.result.emit(result)
            self.signals.log.emit("Excel文件读取完成")
            
        except Exception as e:
            import traceback
            error_msg = f"读取Excel文件失败: {e}"
            self.signals.log.emit(f"✗ {error_msg}")
            self.signals.log.emit(f"详细错误: {traceback.format_exc()[:300]}")
            self.signals.error.emit(error_msg)

    def _read_excel(self, file_path):
        import pandas as pd
        
        with pd.ExcelFile(file_path) as xls:
            sheet_names = list(xls.sheet_names)
        self.signals.log.emit(f"找到 {len(sheet_names)} 个Sheet")
        
        if self.sheet_name and self.sheet_name in sheet_names:
            current_sheet = self.sheet_name
        else:
            current_sheet = sheet_names[0]
        
        headers, data = self._read_sheet_fast(file_path, current_sheet)
        
        return {
            'sheet_names': sheet_names,
            'current_sheet': current_sheet,
            'headers': headers,
            'data': data
        }

    def _read_sheet_fast(self, file_path, sheet_name):
        try:
            import pandas as pd
            
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            df = df.fillna("")
            
            values = df.values.tolist()
            
            if len(values) == 0:
                raise Exception("Sheet为空")
            
            headers = []
            first_row = values[0]
            for i, val in enumerate(first_row):
                if val == "" or (hasattr(val, 'strip') and val.strip() == ""):
                    headers.append(f"列{i + 1}")
                else:
                    headers.append(str(val))
            
            self.signals.log.emit(f"表头: {headers}")
            return headers, values
            
        except Exception as e:
            self.signals.log.emit(f"✗ pandas读取失败: {e}")
            self.signals.log.emit("尝试使用openpyxl直接读取...")
            
            try:
                from openpyxl import load_workbook

                wb = load_workbook(file_path, read_only=True, data_only=True)
                try:
                    ws = wb[sheet_name]
                    data = []
                    for row in ws.iter_rows(values_only=True):
                        # 保留原始 cell 类型（datetime.time/datetime 等），
                        # 由 TableProcessor._format_cell_value 统一规范化显示
                        data.append([
                            cell if cell is not None else ""
                            for cell in row
                        ])
                finally:
                    wb.close()

                if len(data) > 0:
                    headers = [str(h) if h is not None else "" for h in data[0]]
                    headers = [h if h.strip() != "" else f"列{i+1}" for i, h in enumerate(headers)]
                    self.signals.log.emit(f"表头(openpyxl): {headers}")
                    return headers, data
                else:
                    raise Exception("Sheet为空")

            except Exception as e2:
                raise Exception(f"openpyxl读取也失败: {e2}")


class SendWorker(QThread):
    def __init__(
        self,
        tasks,
        send_interval=2,
        chat_delay=0.8,
        send_mode="text",
        minimize_after=True,
        attachment=""
    ):
        super().__init__()
        self.tasks = tasks
        self.send_interval = send_interval
        self.chat_delay = chat_delay
        self.send_mode = send_mode
        self.minimize_after = bool(minimize_after)
        self.attachment = str(attachment or "").strip()
        self.signals = WorkerSignals()
        self.paused_event = threading.Event()
        self.stopped_event = threading.Event()

    def _normalize_task(self, task):
        if isinstance(task, dict):
            if "pending_steps" not in task:
                task["pending_steps"] = self._create_steps(task.get("custom_msg", ""))
            return task

        name, person_data, table_data, recipient, custom_msg = task
        return {
            "name": name,
            "person_data": person_data,
            "table_data": table_data,
            "recipient": recipient,
            "custom_msg": custom_msg,
            "pending_steps": self._create_steps(custom_msg)
        }

    def _create_steps(self, custom_msg):
        steps = []
        if self.send_mode in ("image", "image_text"):
            steps.append({"type": "image", "index": 0})
        if self.send_mode in ("text", "image_text"):
            steps.append({"type": "text", "index": 0})
        if custom_msg:
            steps.append({"type": "custom", "index": 0})
        if self.attachment:
            steps.append({"type": "attachment", "index": 0})
        return steps

    def run(self):
        success_count = 0
        failed_tasks = []
        total_count = len(self.tasks)
        sender = None
        result_emitted = False

        # 附加文件预检：配置了附件但文件不存在时提示并按无附件继续。
        # 必须在 _normalize_task(会读取 self.attachment 生成发送步骤)之前执行。
        if self.attachment and not os.path.exists(self.attachment):
            self.signals.log.emit(
                f"⚠ 附加文件不存在，本次发送不带附件: {self.attachment}"
            )
            self.attachment = ""

        try:
            self.tasks = [self._normalize_task(task) for task in self.tasks]
            self.signals.log.emit("初始化微信客户端...")
            sender = WeChatSender.shared_instance()
            sender.cleanup_temp_images()
            if not sender.initialize():
                # 单例失败不影响下次，清空让后续尝试重新 new WeChat
                WeChatSender.reset_shared_instance()
                self.signals.error.emit("微信未登录或未打开")
                failed_tasks = list(self.tasks)
                self.signals.result.emit((0, total_count, total_count, failed_tasks))
                result_emitted = True
                return
            self.signals.log.emit("微信客户端初始化成功")
            # 让 sender.log 通过信号线程安全地回写到 GUI 日志，同时也落 logging 便于 app.log 排查
            # 保存原始 log 方法，在 finally 中还原，避免共享单例的 log 指向已销毁的 QThread 信号
            _orig_sender_log = sender.log
            try:
                import logging as _logging
                _py_logger = _logging.getLogger("wechat_sender")
                def _send_log_hook(m, _cb=self.signals.log.emit, _lg=_py_logger.info):
                    try:
                        _lg(str(m))
                    except Exception:
                        pass
                    try:
                        _cb(str(m))
                    except Exception:
                        pass
                sender.log = _send_log_hook
            except Exception:
                pass
            
            for i, task in enumerate(self.tasks):
                if self.stopped_event.is_set():
                    self.signals.log.emit("⏹ 发送已停止")
                    failed_tasks.extend(self.tasks[i:])
                    break
                
                while self.paused_event.is_set() and not self.stopped_event.is_set():
                    self.stopped_event.wait(0.1)
                if self.stopped_event.is_set():
                    self.signals.log.emit("⏹ 发送已停止")
                    failed_tasks.extend(self.tasks[i:])
                    break
                
                name = task["name"]
                person_data = task["person_data"]
                table_data = task["table_data"]
                recipient = task["recipient"]
                
                try:
                    if i == 0:
                        self.signals.log.emit("等待微信就绪...")
                        if self.stopped_event.wait(0.2):
                            failed_tasks.extend(self.tasks[i:])
                            break
                    
                    self.signals.log.emit(f"[{i+1}/{total_count}] 正在发送给 {recipient} ({name})...")
                    self.signals.progress.emit(int((i + 1) / total_count * 100))
                    
                    messages = []
                    max_message_length = 2000
                    for j in range(0, len(person_data), max_message_length):
                        messages.append(person_data[j:j+max_message_length])
                    
                    while task["pending_steps"] and not self.stopped_event.is_set():
                        step = task["pending_steps"][0]
                        step_type = step["type"]
                        start_index = step.get("index", 0)

                        if step_type == "image":
                            success, next_index = sender.send_table_images_progress(
                                table_data,
                                recipient,
                                chat_delay=self.chat_delay,
                                start_index=start_index,
                                stop_event=self.stopped_event
                            )
                            step["index"] = next_index
                        elif step_type == "text":
                            success, next_index = sender.send_multiple_messages_progress(
                                messages,
                                recipient,
                                chat_delay=self.chat_delay,
                                start_index=start_index,
                                stop_event=self.stopped_event
                            )
                            step["index"] = next_index
                        elif step_type == "attachment":
                            success = sender.send_file(
                                self.attachment,
                                recipient,
                                chat_delay=self.chat_delay,
                                stop_event=self.stopped_event
                            )
                            if success:
                                self.signals.log.emit(
                                    f"[{i+1}/{total_count}] 已发送附加文件"
                                )
                        else:
                            success = sender.send_message(
                                task["custom_msg"],
                                recipient,
                                chat_delay=self.chat_delay,
                                fast_mode=True,
                                stop_event=self.stopped_event
                            )
                            if success:
                                self.signals.log.emit(
                                    f"[{i+1}/{total_count}] 已发送自定义消息"
                                )

                        if not success:
                            break

                        task["pending_steps"].pop(0)
                        if (
                            task["pending_steps"]
                            and self.stopped_event.wait(0.1)
                        ):
                            break

                    if not task["pending_steps"]:
                        self.signals.log.emit(f"[{i+1}/{total_count}] ✅ 成功发送给 {recipient}")
                        success_count += 1
                    else:
                        if self.stopped_event.is_set():
                            self.signals.log.emit(
                                f"[{i+1}/{total_count}] ⏹ 已停止，保留未完成任务: {recipient}"
                            )
                        else:
                            self.signals.log.emit(f"[{i+1}/{total_count}] ❌ 发送失败: {recipient}")
                        failed_tasks.append(task)
                    
                except Exception as e:
                    self.signals.log.emit(f"[{i+1}/{total_count}] ❌ 发送异常: {name} - {str(e)}")
                    failed_tasks.append(task)
                
                if self.stopped_event.is_set():
                    failed_tasks.extend(self.tasks[i + 1:])
                    self.signals.log.emit("⏹ 发送已停止，剩余任务已保留")
                    break

                if (
                    i < total_count - 1
                    and self.stopped_event.wait(self.send_interval)
                ):
                    failed_tasks.extend(self.tasks[i + 1:])
                    self.signals.log.emit("⏹ 发送已停止，剩余任务已保留")
                    break
            
            failed_count = len(failed_tasks)
            if failed_tasks:
                self.signals.log.emit(f"\n--- 发送失败列表 ({failed_count}人) ---")
                for failed_task in failed_tasks:
                    self.signals.log.emit(
                        f"❌ {failed_task['name']} → {failed_task['recipient']}"
                    )
            
            self.signals.log.emit(f"\n发送完成！成功: {success_count}, 失败: {failed_count}, 总计: {total_count}")
            self.signals.result.emit((success_count, failed_count, total_count, failed_tasks))
            result_emitted = True
            
        except Exception as e:
            import traceback
            self.signals.log.emit(f"❌ 发送线程异常: {str(e)}")
            self.signals.log.emit(f"详细错误: {traceback.format_exc()[:300]}")
            self.signals.error.emit(str(e))
            if not result_emitted:
                remaining_tasks = [
                    task for task in self.tasks
                    if (
                        not isinstance(task, dict)
                        or task.get("pending_steps")
                    )
                ]
                self.signals.result.emit(
                    (success_count, len(remaining_tasks), total_count, remaining_tasks)
                )
        finally:
            if sender:
                # 整个发送任务结束后统一最小化微信窗口（隐私保护，任务级一次）
                if self.minimize_after:
                    try:
                        sender.minimize_window()
                    except Exception:
                        pass
                # 还原 sender.log，避免共享单例的 log 回调指向已销毁的 QThread 信号
                try:
                    sender.log = _orig_sender_log
                except Exception:
                    pass
                sender.cleanup_temp_images()

    def set_paused(self, value):
        if value:
            self.paused_event.set()
        else:
            self.paused_event.clear()

    def set_stopped(self, value):
        if value:
            self.stopped_event.set()
        else:
            self.stopped_event.clear()

    def is_paused(self):
        return self.paused_event.is_set()

    def is_stopped(self):
        return self.stopped_event.is_set()


class _SenderLogCb:
    """把 WeChatSender.log 映射到定时任务的 GUI 日志面板。

    设计成类 + 闭包是因为：
    1) 需要可 pickle 弱引用（sender.log 可能被反复设置），lambda 容易循环引用；
    2) 吞掉所有异常，避免 wxauto4 调用链在 GUI 面板写入异常时被连带中断。
    3) 只写入 GUI 面板，**不再**额外调用 logging。
       WeChatSender 内部会自己再调 logger.info 写 app.log，两者分离、不重复。
    """
    __slots__ = ("_cb",)

    def __init__(self, cb):
        self._cb = cb

    def log(self, message):
        cb = self._cb
        if cb is None:
            return
        try:
            cb(str(message))
        except Exception:
            # 面板写入失败（控件销毁/UI线程异常）只吞掉，不能阻断微信发送链路。
            pass


class ScheduleSendWorker(QThread):
    """定时消息发送 Worker：为多个接收人逐条发送自定义文字，含3次重试+模糊匹配。"""

    finished_with_result = pyqtSignal(int, int, list)  # success, failed, failed_recipients

    def __init__(
        self,
        recipients,
        message,
        chat_delay=0.3,
        send_interval=0.5,
        log_callback=None,
        default_city="",
        minimize_after=True,
        attachment="",
    ):
        super().__init__()
        self.recipients = list(recipients or [])
        self.message = message or ""
        self.chat_delay = max(0.0, min(10.0, float(chat_delay)))
        self.send_interval = max(0.0, min(30.0, float(send_interval)))
        self.default_city = str(default_city or "")
        self.minimize_after = bool(minimize_after)
        self.attachment = str(attachment or "").strip()
        self._log_cb = log_callback
        self.stopped_event = threading.Event()
        self._sender = None
        self._first_send = True

    def stop(self):
        self.stopped_event.set()

    def log(self, msg):
        if self._log_cb:
            try:
                self._log_cb(msg)
            except Exception:
                pass

    def run(self):
        success = 0
        failed_recipients = []
        # 锁屏预警：Windows 安全桌面限制导致程序无法对已锁定的会话自动解锁
        try:
            if is_workstation_locked():
                self.log(
                    "[定时] ⚠ 电脑处于锁定状态且无法自动解锁，本次发送很可能失败；"
                    "请保持电脑解锁或开启'定时期间保持电脑解锁'"
                )
        except Exception:
            pass
        if not self.message.strip():
            self.log("[定时] 消息为空，跳过发送")
            self.finished_with_result.emit(0, len(self.recipients), list(self.recipients))
            return
        if not self.recipients:
            self.log("[定时] 接收人为空，跳过发送")
            self.finished_with_result.emit(0, 0, [])
            return

        # 天气占位符渲染：消息里若含 {{weather}}/{{temp}} 等占位符，调用
        # weather_fetcher 替换为真实值；无 API key 时跳过渲染保留原文
        try:
            from modules import weather_fetcher
            cfg = weather_fetcher.load_weather_config()
            placeholders = weather_fetcher.find_placeholders(self.message)
            if placeholders:
                self.log(f"[定时] 检测到消息占位符: {', '.join(placeholders)}")
                # 不需要天气 key 的占位符（news/date/weekday/time）也会被渲染
                rendered = weather_fetcher.render_message(
                    self.message,
                    task_default_city=self.default_city,
                    api_key=cfg.get("api_key", ""),
                    base_url=cfg.get("base_url", ""),
                    global_default_city=cfg.get("default_city", ""),
                )
                if rendered != self.message:
                    self.log("[定时] 占位符渲染完成")
                    self.message = rendered
                else:
                    self.log("[定时] 占位符未生效（检查天气 API key/城市/网络）")
        except Exception as exc:
            self.log(f"[定时] 天气渲染异常: {exc}")

        self.log("[定时] 初始化微信客户端...")
        _orig_sender_log = None
        try:
            self._sender = WeChatSender.shared_instance()
            # 保存原始 log 方法，在结束时还原，避免共享单例的 log 指向已销毁的 QThread
            _orig_sender_log = self._sender.log
            if self._log_cb:
                self._sender.log = _SenderLogCb(self._log_cb).log
            if not self._sender.initialize():
                WeChatSender.reset_shared_instance()
                self.log("[定时] 微信初始化失败，本次任务失败")
                self.finished_with_result.emit(0, len(self.recipients), list(self.recipients))
                return
            self.log("[定时] 微信客户端初始化成功")
        except Exception as exc:
            self.log(f"[定时] 初始化异常: {exc}")
            WeChatSender.reset_shared_instance()
            self.finished_with_result.emit(0, len(self.recipients), list(self.recipients))
            return

        total = len(self.recipients)
        # 附加文件预检：配置了附件但文件不存在时提示并跳过附件（文字照常发送）
        attachment_ok = False
        if self.attachment:
            if os.path.exists(self.attachment):
                attachment_ok = True
                self.log(f"[定时] 本任务将附加发送文件: {os.path.basename(self.attachment)}")
            else:
                self.log(f"[定时] ⚠ 附加文件不存在，跳过附件发送: {self.attachment}")
        try:
            for idx, recipient in enumerate(self.recipients):
                if self.stopped_event.is_set():
                    self.log("[定时] 已手动停止")
                    failed_recipients.extend(self.recipients[idx:])
                    break
                if idx > 0 and self.send_interval > 0:
                    if self.stopped_event.wait(self.send_interval):
                        failed_recipients.extend(self.recipients[idx:])
                        break
                recipient = str(recipient).strip()
                if not recipient:
                    continue
                self.log(f"[定时] [{idx+1}/{total}] 发送到 {recipient}")
                ok = self._sender.send_message(
                    content=self.message,
                    recipient=recipient,
                    first_send=self._first_send,
                    chat_delay=self.chat_delay,
                    fast_mode=False,
                    stop_event=self.stopped_event,
                )
                self._first_send = False
                # 文字发送成功后再发附加文件；附件失败只记录日志，
                # 不把整个接收人判为失败（避免重试时文字重复发送）
                if ok and attachment_ok:
                    if not self._sender.send_file(
                        self.attachment,
                        recipient,
                        chat_delay=self.chat_delay,
                        stop_event=self.stopped_event,
                    ):
                        self.log(f"[定时] ⚠ 附加文件发送失败: {recipient}")
                if ok:
                    success += 1
                else:
                    failed_recipients.append(recipient)
                    self.log(f"[定时] ✗ 发送失败: {recipient}")
        finally:
            if self._sender is not None:
                # 整个定时任务结束后统一最小化微信窗口（任务级一次，按任务开关）
                if self.minimize_after:
                    try:
                        self._sender.minimize_window()
                    except Exception:
                        pass
                # 还原 sender.log，避免共享单例的 log 回调指向已结束的 worker
                if _orig_sender_log is not None:
                    try:
                        self._sender.log = _orig_sender_log
                    except Exception:
                        pass
        self.finished_with_result.emit(success, len(failed_recipients), failed_recipients)


class TableFilterTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.table_data = None
        self.processor = None
        self.wechat_mapping = {}
        self.filter_conditions = []
        self.current_excel_path = None
        self.sheet_names = []
        self.current_sheet = None
        self.headers = []
        self.excel_worker = None
        self.worker = None
        self.last_failed_tasks = []
        self.config_manager = ConfigManager()
        self.current_config_path = None
        self.pending_config = None
        
        self.init_ui()
        self.connect_signals()
        self.refresh_recent_configs()

    def init_ui(self):
        main_layout = QVBoxLayout()
        
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        config_group = QGroupBox("配置管理")
        config_layout = QVBoxLayout(config_group)

        self.current_config_label = QLabel("当前配置: 未加载")
        self.current_config_label.setWordWrap(True)
        self.current_config_label.setStyleSheet(
            "color: #666; font-size: 11px;"
        )
        config_layout.addWidget(self.current_config_label)

        config_btn_layout = QHBoxLayout()
        self.save_config_btn = QPushButton("保存配置")
        self.save_config_btn.setEnabled(False)
        config_btn_layout.addWidget(self.save_config_btn)

        self.save_as_config_btn = QPushButton("另存为配置")
        self.save_as_config_btn.setEnabled(False)
        config_btn_layout.addWidget(self.save_as_config_btn)

        self.load_config_btn = QPushButton("加载配置")
        config_btn_layout.addWidget(self.load_config_btn)
        config_layout.addLayout(config_btn_layout)

        config_layout.addWidget(QLabel("最近配置（双击加载）:"))
        self.recent_config_list = QListWidget()
        self.recent_config_list.setMinimumHeight(50)
        self.recent_config_list.setMaximumHeight(72)
        config_layout.addWidget(self.recent_config_list)

        left_layout.addWidget(config_group)

        url_group = QGroupBox("文档设置")
        url_layout = QVBoxLayout(url_group)
        
        excel_btn_layout = QHBoxLayout()
        self.excel_btn = QPushButton("📂 选择Excel文件")
        self.excel_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 5px 8px; font-size: 12px;")
        excel_btn_layout.addWidget(self.excel_btn)
        
        self.reload_btn = QPushButton("🔄 重新读取")
        self.reload_btn.setStyleSheet("background-color: #FF9800; color: white; padding: 5px 8px; font-size: 12px;")
        self.reload_btn.setEnabled(False)
        excel_btn_layout.addWidget(self.reload_btn)
        
        self.open_excel_btn = QPushButton("📝 打开文件")
        self.open_excel_btn.setStyleSheet("background-color: #2196F3; color: white; padding: 5px 8px; font-size: 12px;")
        self.open_excel_btn.setEnabled(False)
        excel_btn_layout.addWidget(self.open_excel_btn)
        
        url_layout.addLayout(excel_btn_layout)
        
        self.current_file_label = QLabel("当前文件: 未选择")
        self.current_file_label.setWordWrap(True)
        self.current_file_label.setStyleSheet("color: #666; font-size: 12px;")
        url_layout.addWidget(self.current_file_label)
        
        url_layout.addWidget(QLabel("Sheet名称:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setPlaceholderText("请先选择Excel文件")
        self.sheet_combo.setEnabled(False)
        url_layout.addWidget(self.sheet_combo)
        
        url_layout.addWidget(QLabel("人名所在列:"))
        self.name_column_combo = QComboBox()
        self.name_column_combo.setPlaceholderText("请先选择Excel文件")
        self.name_column_combo.setEnabled(False)
        url_layout.addWidget(self.name_column_combo)
        
        url_layout.addWidget(QLabel("要提取的列(多选):"))
        self.extract_columns_btn = QPushButton("点击选择要提取的列")
        self.extract_columns_btn.setStyleSheet("background-color: #f0f0f0; color: #333; padding: 5px 8px; font-size: 12px;")
        self.extract_columns_btn.setEnabled(False)
        url_layout.addWidget(self.extract_columns_btn)
        
        self.selected_columns_label = QLabel("已选择: 0列")
        self.selected_columns_label.setStyleSheet("color: #666; font-size: 11px;")
        url_layout.addWidget(self.selected_columns_label)
        
        self.selected_columns = []
        
        url_layout.addWidget(QLabel("微信昵称所在列(可选):"))
        self.wechat_column_combo = QComboBox()
        self.wechat_column_combo.setPlaceholderText("请先选择Excel文件")
        self.wechat_column_combo.setEnabled(False)
        url_layout.addWidget(self.wechat_column_combo)
        
        self.load_data_btn = QPushButton("🚀 加载数据")
        self.load_data_btn.setStyleSheet("background-color: #9C27B0; color: white; padding: 6px; font-size: 12px; font-weight: bold;")
        self.load_data_btn.setEnabled(False)
        url_layout.addWidget(self.load_data_btn)
        
        left_layout.addWidget(url_group)
        
        filter_group = QGroupBox("筛选条件")
        filter_layout = QVBoxLayout(filter_group)
        
        self.filter_conditions_layout = QVBoxLayout()
        filter_layout.addLayout(self.filter_conditions_layout)
        
        filter_btn_layout = QHBoxLayout()
        self.add_filter_btn = QPushButton("+ 添加条件")
        self.add_filter_btn.setStyleSheet("background-color: #FF9800; color: white; padding: 4px;")
        filter_btn_layout.addWidget(self.add_filter_btn)
        
        self.clear_filter_btn = QPushButton("清除条件")
        self.clear_filter_btn.setStyleSheet("background-color: #f44336; color: white; padding: 4px;")
        filter_btn_layout.addWidget(self.clear_filter_btn)
        
        filter_layout.addLayout(filter_btn_layout)
        
        self.apply_filter_btn = QPushButton("应用筛选")
        self.apply_filter_btn.setStyleSheet("background-color: #00BCD4; color: white; padding: 4px;")
        self.apply_filter_btn.setEnabled(False)
        filter_layout.addWidget(self.apply_filter_btn)
        
        left_layout.addWidget(filter_group)
        
        left_layout.addStretch()
        splitter.addWidget(left_panel)
        
        middle_panel = QWidget()
        middle_layout = QVBoxLayout(middle_panel)
        
        persons_group = QGroupBox("人员列表")
        persons_layout = QVBoxLayout(persons_group)
        
        self.persons_list = QListWidget()
        self.persons_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.persons_list.setMinimumHeight(150)
        persons_layout.addWidget(self.persons_list)
        
        btn_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("全选")
        self.deselect_all_btn = QPushButton("取消全选")
        btn_layout.addWidget(self.select_all_btn)
        btn_layout.addWidget(self.deselect_all_btn)
        persons_layout.addLayout(btn_layout)
        
        middle_layout.addWidget(persons_group)
        
        preview_group = QGroupBox("数据预览")
        preview_layout = QVBoxLayout(preview_group)
        
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMinimumHeight(150)
        self.preview_text.setMaximumHeight(200)
        preview_layout.addWidget(self.preview_text)
        
        middle_layout.addWidget(preview_group)
        
        middle_layout.addStretch()
        splitter.addWidget(middle_panel)
        
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        send_group = QGroupBox("发送设置")
        send_layout = QVBoxLayout(send_group)
        
        send_layout.addWidget(QLabel("微信接收人(手动指定):"))
        self.wechat_edit = QLineEdit()
        send_layout.addWidget(self.wechat_edit)
        
        send_mode_layout = QHBoxLayout()
        send_mode_layout.addWidget(QLabel("数据发送形式:"))
        self.send_mode_combo = QComboBox()
        self.send_mode_combo.addItem("文字", "text")
        self.send_mode_combo.addItem("图片", "image")
        self.send_mode_combo.addItem("图片后再发文字", "image_text")
        send_mode_layout.addWidget(self.send_mode_combo)
        send_layout.addLayout(send_mode_layout)

        interval_layout = QHBoxLayout()
        interval_layout.addWidget(QLabel("发送间隔(秒):"))
        self.send_interval_spin = QDoubleSpinBox()
        self.send_interval_spin.setRange(0.0, 30.0)
        self.send_interval_spin.setSingleStep(0.1)
        self.send_interval_spin.setValue(0.5)
        interval_layout.addWidget(self.send_interval_spin)
        send_layout.addLayout(interval_layout)
        
        chat_delay_layout = QHBoxLayout()
        chat_delay_layout.addWidget(QLabel("聊天窗口延迟(秒):"))
        self.chat_delay_spin = QDoubleSpinBox()
        self.chat_delay_spin.setRange(0.0, 10.0)
        self.chat_delay_spin.setSingleStep(0.1)
        self.chat_delay_spin.setValue(0.3)
        self.chat_delay_spin.setFixedWidth(100)
        chat_delay_layout.addWidget(self.chat_delay_spin)
        send_layout.addLayout(chat_delay_layout)
        
        self.custom_msg_checkbox = QCheckBox("发送后追加自定义消息")
        send_layout.addWidget(self.custom_msg_checkbox)
        
        send_layout.addWidget(QLabel("自定义消息内容:"))
        self.custom_msg_edit = QTextEdit()
        self.custom_msg_edit.setPlaceholderText("发送完表格数据后，会额外发送这条消息...")
        self.custom_msg_edit.setMaximumHeight(60)
        self.custom_msg_edit.setEnabled(False)
        send_layout.addWidget(self.custom_msg_edit)

        attach_row = QHBoxLayout()
        attach_row.addWidget(QLabel("附加文件:"))
        self.attachment_edit = QLineEdit()
        self.attachment_edit.setReadOnly(True)
        self.attachment_edit.setPlaceholderText("可选，选择文件/图片随消息一起发送")
        attach_row.addWidget(self.attachment_edit, 1)
        self.attach_pick_btn = QPushButton("选择...")
        self.attach_clear_btn = QPushButton("清除")
        self.attach_pick_btn.setFixedWidth(60)
        self.attach_clear_btn.setFixedWidth(50)
        attach_row.addWidget(self.attach_pick_btn)
        attach_row.addWidget(self.attach_clear_btn)
        send_layout.addLayout(attach_row)
        
        self.send_btn = QPushButton("发送选中人员数据")
        self.send_btn.setStyleSheet("background-color: #2196F3; color: white; padding: 5px 8px; font-size: 12px;")
        self.send_btn.setEnabled(False)
        send_layout.addWidget(self.send_btn)
        
        right_layout.addWidget(send_group)
        
        progress_group = QGroupBox("发送进度")
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        self.progress_label = QLabel("等待发送...")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        progress_layout.addWidget(self.progress_label)
        
        right_layout.addWidget(progress_group)
        
        log_group = QGroupBox("日志")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_text)
        
        right_layout.addWidget(log_group)
        
        control_group = QGroupBox("发送控制")
        control_layout = QHBoxLayout(control_group)
        
        self.start_send_btn = QPushButton("▶ 开始发送")
        self.start_send_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 8px; font-size: 14px;")
        self.start_send_btn.setEnabled(False)
        control_layout.addWidget(self.start_send_btn)
        
        self.pause_send_btn = QPushButton("⏸ 暂停")
        self.pause_send_btn.setStyleSheet("background-color: #FF9800; color: white; padding: 8px; font-size: 14px;")
        self.pause_send_btn.setEnabled(False)
        control_layout.addWidget(self.pause_send_btn)
        
        self.stop_send_btn = QPushButton("⏹ 停止")
        self.stop_send_btn.setStyleSheet("background-color: #f44336; color: white; padding: 8px; font-size: 14px;")
        self.stop_send_btn.setEnabled(False)
        control_layout.addWidget(self.stop_send_btn)
        
        self.retry_send_btn = QPushButton("🔄 重试发送")
        self.retry_send_btn.setStyleSheet("background-color: #00BCD4; color: white; padding: 8px; font-size: 14px;")
        self.retry_send_btn.setEnabled(False)
        control_layout.addWidget(self.retry_send_btn)

        self.minimize_check = QCheckBox("发送后最小化微信")
        self.minimize_check.setToolTip(
            "开启后：整个发送任务完成后把微信窗口最小化一次（保护隐私）。\n"
            "关闭后微信窗口保持原样，不会被最小化。"
        )
        self.minimize_check.setChecked(True)
        control_layout.addWidget(self.minimize_check)

        control_layout.addStretch()
        
        right_layout.addWidget(control_group)
        
        right_layout.addStretch()
        splitter.addWidget(right_panel)
        
        splitter.setSizes([285, 280, 285])
        
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

    def connect_signals(self):
        self.save_config_btn.clicked.connect(self.save_config)
        self.save_as_config_btn.clicked.connect(self.save_config_as)
        self.load_config_btn.clicked.connect(self.load_config)
        self.recent_config_list.itemDoubleClicked.connect(
            self.load_recent_config
        )
        self.excel_btn.clicked.connect(self.select_excel_file)
        self.reload_btn.clicked.connect(self.reload_excel_file)
        self.open_excel_btn.clicked.connect(self.open_current_excel)
        self.sheet_combo.currentIndexChanged.connect(self.on_sheet_changed)
        self.extract_columns_btn.clicked.connect(self.open_extract_columns_dialog)
        self.load_data_btn.clicked.connect(self.load_data)
        self.start_send_btn.clicked.connect(self.send_data)
        self.pause_send_btn.clicked.connect(self.pause_send)
        self.stop_send_btn.clicked.connect(self.stop_send)
        self.retry_send_btn.clicked.connect(self.retry_send)
        self.send_btn.clicked.connect(self.send_data)
        self.select_all_btn.clicked.connect(lambda: self.persons_list.selectAll())
        self.deselect_all_btn.clicked.connect(lambda: self.persons_list.clearSelection())
        self.persons_list.itemSelectionChanged.connect(self.on_person_selection)
        self.add_filter_btn.clicked.connect(self.add_filter_condition)
        self.clear_filter_btn.clicked.connect(self.clear_filter_conditions)
        self.apply_filter_btn.clicked.connect(self.apply_filter)
        self.custom_msg_checkbox.stateChanged.connect(self.on_custom_msg_checkbox_changed)
        self.attach_pick_btn.clicked.connect(self._on_pick_attachment)
        self.attach_clear_btn.clicked.connect(self._on_clear_attachment)

    def refresh_recent_configs(self):
        self.recent_config_list.clear()
        for config_path in self.config_manager.get_recent_profiles():
            item = QListWidgetItem(os.path.basename(config_path))
            item.setToolTip(config_path)
            item.setData(Qt.ItemDataRole.UserRole, config_path)
            self.recent_config_list.addItem(item)

    def _set_current_config(self, config_path):
        self.current_config_path = os.path.abspath(config_path)
        self.current_config_label.setText(
            f"当前配置: {os.path.basename(self.current_config_path)}"
        )
        self.current_config_label.setToolTip(self.current_config_path)

    def _build_profile_data(self):
        if self.excel_worker and self.excel_worker.isRunning():
            QMessageBox.warning(self, "警告", "Excel 正在读取，请稍候再保存配置")
            return None

        if not self.current_excel_path or not self.headers:
            QMessageBox.warning(self, "警告", "请先读取 Excel 文件")
            return None

        name_column = self.name_column_combo.currentText().strip()
        if not name_column:
            QMessageBox.warning(self, "警告", "请选择人员所在列")
            return None

        if not self.selected_columns:
            QMessageBox.warning(self, "警告", "请至少选择一列提取信息")
            return None

        wechat_column = ""
        if self.wechat_column_combo.currentIndex() > 0:
            wechat_column = self.wechat_column_combo.currentText().strip()

        return {
            "version": ConfigManager.PROFILE_VERSION,
            "excel": {
                "path": os.path.abspath(self.current_excel_path),
                "sheet": self.sheet_combo.currentText().strip(),
                "name_column": name_column,
                "extract_columns": list(self.selected_columns),
                "wechat_column": wechat_column,
            },
            "send": {
                "manual_recipient": self.wechat_edit.text().strip(),
                "mode": self.send_mode_combo.currentData(),
                "interval": self.send_interval_spin.value(),
                "chat_delay": self.chat_delay_spin.value(),
                "custom_message_enabled": (
                    self.custom_msg_checkbox.isChecked()
                ),
                "custom_message": self.custom_msg_edit.toPlainText(),
                "attachment": self.attachment_edit.text().strip(),
            },
        }

    def save_config(self):
        if not self.current_config_path:
            self.save_config_as()
            return
        self._save_config_to(self.current_config_path)

    def save_config_as(self):
        profile = self._build_profile_data()
        if not profile:
            return

        if self.current_config_path:
            current_name = os.path.splitext(
                os.path.basename(self.current_config_path)
            )[0]
            default_path = os.path.join(
                os.path.dirname(self.current_config_path),
                f"{current_name}_副本.json",
            )
        else:
            excel_name = os.path.splitext(
                os.path.basename(self.current_excel_path)
            )[0]
            default_path = os.path.join(
                str(self.config_manager.profile_dir),
                f"{excel_name}.json",
            )

        config_path, _ = QFileDialog.getSaveFileName(
            self,
            "另存为配置",
            default_path,
            "表格发送配置 (*.json)",
        )
        if config_path:
            self._save_config_to(config_path, profile)

    def _save_config_to(self, config_path, profile=None):
        profile = profile or self._build_profile_data()
        if not profile:
            return False

        try:
            saved_path = self.config_manager.save_profile(
                config_path,
                profile,
            )
        except ConfigError as error:
            QMessageBox.warning(self, "保存失败", str(error))
            self.log(f"✗ 配置保存失败: {error}")
            return False

        self._set_current_config(saved_path)
        self.refresh_recent_configs()
        self.log(f"配置已保存: {saved_path}")
        QMessageBox.information(
            self,
            "保存成功",
            f"配置已保存到:\n{saved_path}",
        )
        return True

    def load_config(self):
        if not self._can_load_config():
            return

        start_path = (
            os.path.dirname(self.current_config_path)
            if self.current_config_path
            else str(self.config_manager.profile_dir)
        )
        config_path, _ = QFileDialog.getOpenFileName(
            self,
            "加载配置",
            start_path,
            "表格发送配置 (*.json)",
        )
        if config_path:
            self._load_config_file(config_path)

    def load_recent_config(self, item):
        config_path = item.data(Qt.ItemDataRole.UserRole)
        if not config_path:
            return
        if not os.path.isfile(config_path):
            self.config_manager.remove_recent(config_path)
            self.refresh_recent_configs()
            QMessageBox.warning(self, "配置不存在", "该配置文件已被移动或删除")
            return
        self._load_config_file(config_path)

    def _can_load_config(self):
        if self.excel_worker and self.excel_worker.isRunning():
            QMessageBox.warning(self, "请稍候", "Excel 正在读取，暂时不能加载配置")
            return False
        if self.worker and self.worker.isRunning():
            QMessageBox.warning(self, "请稍候", "正在发送消息，暂时不能加载配置")
            return False
        return True

    def _load_config_file(self, config_path):
        if not self._can_load_config():
            return False

        try:
            profile = self.config_manager.load_profile(config_path)
        except ConfigError as error:
            QMessageBox.warning(self, "加载失败", str(error))
            self.log(f"✗ 配置加载失败: {error}")
            return False

        excel_path = profile["excel"]["path"]
        if not os.path.isfile(excel_path):
            QMessageBox.warning(
                self,
                "Excel 文件不存在",
                "配置中的 Excel 文件已被移动或删除，请重新选择该文件。",
            )
            excel_path, _ = QFileDialog.getOpenFileName(
                self,
                "重新定位 Excel 文件",
                os.path.dirname(excel_path),
                "Excel文件 (*.xlsx *.xls)",
            )
            if not excel_path:
                return False
            profile["excel"]["path"] = os.path.abspath(excel_path)

        self._set_current_config(config_path)
        self.config_manager.add_recent(self.current_config_path)
        self.refresh_recent_configs()

        self.pending_config = {
            "profile": profile,
            "excel_path": os.path.abspath(excel_path),
        }
        sheet_name = profile["excel"].get("sheet") or None
        if not self._load_excel_data(excel_path, sheet_name):
            self.pending_config = None
            return False

        self.log(
            f"正在加载配置: {os.path.basename(self.current_config_path)}"
        )
        return True

    def _apply_pending_config(self, worker):
        if not self.pending_config:
            return False

        expected_path = os.path.normcase(
            os.path.abspath(self.pending_config["excel_path"])
        )
        actual_path = os.path.normcase(os.path.abspath(worker.file_path))
        if expected_path != actual_path:
            return False

        profile = self.pending_config["profile"]
        self.pending_config = None
        excel_settings = profile["excel"]
        send_settings = profile["send"]
        missing_settings = []

        saved_sheet = excel_settings.get("sheet", "")
        if saved_sheet and saved_sheet != self.current_sheet:
            missing_settings.append(
                f"Sheet“{saved_sheet}”不存在，已使用“{self.current_sheet}”"
            )

        name_column = excel_settings["name_column"]
        if name_column in self.headers:
            self.name_column_combo.setCurrentText(name_column)
        else:
            missing_settings.append(f"人员列“{name_column}”不存在")

        selected_columns = [
            column
            for column in excel_settings["extract_columns"]
            if column in self.headers
        ]
        missing_columns = [
            column
            for column in excel_settings["extract_columns"]
            if column not in self.headers
        ]
        self.selected_columns = selected_columns
        self._update_selected_columns_label()
        if missing_columns:
            missing_settings.append(
                f"提取列不存在: {', '.join(missing_columns)}"
            )

        wechat_column = excel_settings.get("wechat_column", "")
        if not wechat_column:
            self.wechat_column_combo.setCurrentIndex(0)
        elif wechat_column in self.headers:
            self.wechat_column_combo.setCurrentText(wechat_column)
        else:
            self.wechat_column_combo.setCurrentIndex(0)
            missing_settings.append(f"微信列“{wechat_column}”不存在")

        self.wechat_edit.setText(
            send_settings.get("manual_recipient", "")
        )
        send_mode_index = self.send_mode_combo.findData(
            send_settings.get("mode", "text")
        )
        if send_mode_index >= 0:
            self.send_mode_combo.setCurrentIndex(send_mode_index)
        self.send_interval_spin.setValue(send_settings.get("interval", 0.5))
        self.chat_delay_spin.setValue(
            send_settings.get("chat_delay", 0.3)
        )
        self.custom_msg_checkbox.setChecked(
            send_settings.get("custom_message_enabled", False)
        )
        self.custom_msg_edit.setPlainText(
            send_settings.get("custom_message", "")
        )
        self.attachment_edit.setText(
            send_settings.get("attachment", "") or ""
        )

        can_load_data = (
            name_column in self.headers
            and bool(self.selected_columns)
        )
        if can_load_data:
            self.load_data()
            self.log("配置已应用，并已自动加载人员数据")
        else:
            self.log("⚠ 配置已部分应用，请重新选择缺失的列")

        if missing_settings:
            QMessageBox.warning(
                self,
                "配置部分失效",
                "\n".join(missing_settings),
            )
        return True

    def get_available_operators(self):
        return [
            ("等于", "equals"),
            ("包含", "contains"),
            ("不包含", "not_contains"),
            ("为空", "empty"),
            ("不为空", "not_empty"),
            ("大于", "greater_than"),
            ("小于", "less_than"),
            ("大于等于", "equals_or_greater"),
            ("小于等于", "equals_or_less"),
            ("日期等于", "date_equals"),
            ("日期之后", "date_after"),
            ("日期之前", "date_before")
        ]

    def add_filter_condition(self):
        if not self.processor:
            QMessageBox.warning(self, "警告", "请先提取表格数据")
            return
        
        headers = self.processor.get_headers()
        
        row_widget = QWidget()
        row_layout = QHBoxLayout(row_widget)
        row_layout.setContentsMargins(2, 2, 2, 2)
        
        col_combo = QComboBox()
        col_combo.addItems(headers)
        row_layout.addWidget(col_combo)
        
        op_combo = QComboBox()
        op_names = [name for name, _ in self.get_available_operators()]
        op_combo.addItems(op_names)
        row_layout.addWidget(op_combo)
        
        value_edit = QLineEdit()
        value_edit.setPlaceholderText("条件值")
        row_layout.addWidget(value_edit)
        
        remove_btn = QPushButton("×")
        remove_btn.setStyleSheet("background-color: #f44336; color: white; padding: 2px 6px;")
        row_layout.addWidget(remove_btn)
        
        remove_btn.clicked.connect(lambda: self.remove_filter_condition(row_widget))
        
        self.filter_conditions_layout.addWidget(row_widget)
        self.apply_filter_btn.setEnabled(True)
        
        col_combo.currentTextChanged.connect(self.on_filter_change)
        op_combo.currentTextChanged.connect(self.on_filter_change)
        value_edit.textChanged.connect(self.on_filter_change)

    def remove_filter_condition(self, widget):
        self.filter_conditions_layout.removeWidget(widget)
        widget.deleteLater()
        
        if self.filter_conditions_layout.count() == 0:
            self.apply_filter_btn.setEnabled(False)
            self.filter_conditions = []

    def clear_filter_conditions(self):
        while self.filter_conditions_layout.count() > 0:
            item = self.filter_conditions_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        self.apply_filter_btn.setEnabled(False)
        self.filter_conditions = []
        
        if self.processor:
            self.refresh_persons_list()

    def on_filter_change(self):
        pass

    def on_custom_msg_checkbox_changed(self, state):
        self.custom_msg_edit.setEnabled(state == Qt.CheckState.Checked.value)

    def _on_pick_attachment(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择附加文件(文件/图片)",
            "",
            "常用文件 (*.png *.jpg *.jpeg *.gif *.bmp *.webp *.pdf *.doc *.docx *.xls *.xlsx *.ppt *.pptx *.txt *.zip *.rar *.7z);;所有文件 (*.*)",
        )
        if path:
            self.attachment_edit.setText(path)

    def _on_clear_attachment(self):
        self.attachment_edit.clear()

    def open_extract_columns_dialog(self):
        if not hasattr(self, 'headers') or not self.headers:
            return
        
        dialog = MultiSelectDialog("选择要提取的列", self.headers, self.selected_columns, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.selected_columns = dialog.selected_columns
            self._update_selected_columns_label()

    def _update_selected_columns_label(self):
        if not self.selected_columns:
            self.selected_columns_label.setText("已选择: 0列")
        elif len(self.selected_columns) <= 3:
            self.selected_columns_label.setText(
                f"已选择: {', '.join(self.selected_columns)}"
            )
        else:
            self.selected_columns_label.setText(
                f"已选择: {len(self.selected_columns)}列"
            )

    def build_filter_conditions(self):
        conditions = []
        from modules.table_processor import FilterCondition
        
        for i in range(self.filter_conditions_layout.count()):
            item = self.filter_conditions_layout.itemAt(i)
            if item and item.widget():
                widget = item.widget()
                layout = widget.layout()
                
                col_combo = layout.itemAt(0).widget()
                op_combo = layout.itemAt(1).widget()
                value_edit = layout.itemAt(2).widget()
                
                column_name = col_combo.currentText()
                op_index = op_combo.currentIndex()
                operator = self.get_available_operators()[op_index][1]
                value = value_edit.text().strip()
                
                conditions.append(FilterCondition(column_name, operator, value))
        
        return conditions

    def apply_filter(self):
        if not self.processor:
            return
        
        self.filter_conditions = self.build_filter_conditions()
        self.refresh_persons_list()
        
        condition_descriptions = []
        for cond in self.filter_conditions:
            op_name = dict(self.get_available_operators()).get(cond.operator, cond.operator)
            if cond.operator in ["empty", "not_empty"]:
                condition_descriptions.append(f"{cond.column_name} {op_name}")
            else:
                condition_descriptions.append(f"{cond.column_name} {op_name} '{cond.value}'")
        
        if condition_descriptions:
            self.log(f"应用筛选条件: {'，'.join(condition_descriptions)}")
        else:
            self.log("清除筛选条件，显示全部人员")

    def refresh_persons_list(self):
        if not self.processor:
            return
        
        name_column = self.name_column_combo.currentText()
        persons = self.processor.get_all_persons(name_column, self.filter_conditions)
        
        self.persons_list.clear()
        for person in persons:
            self.persons_list.addItem(person)

        # 列表重建会丢失选中项：只要还有人员就保持"开始发送"可点击，
        # 未选中人员时点击会弹提示，避免禁用按钮导致"点击毫无反应"
        self.start_send_btn.setEnabled(self.persons_list.count() > 0)

        self.log(f"筛选后找到 {len(persons)} 个人")

    def on_person_selection(self):
        selected_items = self.persons_list.selectedItems()
        if selected_items and self.table_data:
            self.send_btn.setEnabled(True)
            self.start_send_btn.setEnabled(True)
            self.preview_selected_data()
        else:
            self.send_btn.setEnabled(False)
            # 列表还有人员时保持"开始发送"可点击（未选中时点击会弹提示）
            self.start_send_btn.setEnabled(self.persons_list.count() > 0)

    def preview_selected_data(self):
        if not self.processor:
            return
        
        selected_items = self.persons_list.selectedItems()
        if not selected_items:
            return
        
        extract_columns = self.selected_columns
        if not extract_columns:
            extract_columns = [self.headers[0]] if self.headers else []
        
        preview_content = []
        
        for item in selected_items:
            name = item.text()
            person_data = self.processor.get_person_data(
                name, 
                self.name_column_combo.currentText(), 
                extract_columns,
                self.filter_conditions
            )
            if person_data:
                preview_content.append(f"=== {name} ===")
                preview_content.append(person_data)
                preview_content.append("")
        
        self.preview_text.setPlainText("\n".join(preview_content))

    def select_excel_file(self):
        if self.excel_worker and self.excel_worker.isRunning():
            self.log("Excel文件正在读取，请稍候...")
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls)"
        )
        
        if file_path:
            self.pending_config = None
            self._load_excel_data(file_path)
    
    def reload_excel_file(self):
        if not getattr(self, 'current_excel_path', None):
            QMessageBox.warning(self, "警告", "请先选择Excel文件")
            return

        if self.excel_worker and self.excel_worker.isRunning():
            self.log("Excel文件正在读取，请稍候...")
            return

        self.log("重新读取Excel文件...")
        self._load_excel_data(self.current_excel_path)
    
    def open_current_excel(self):
        if hasattr(self, 'current_excel_path') and self.current_excel_path:
            try:
                os.startfile(self.current_excel_path)
                self.log(f"已打开文件: {self.current_excel_path}")
            except Exception as e:
                self.log(f"✗ 打开文件失败: {e}")
                QMessageBox.warning(self, "打开失败", f"无法打开文件: {e}")
    
    def _set_excel_loading(self, loading):
        self.excel_btn.setEnabled(not loading)
        has_file = bool(self.current_excel_path)
        has_data = bool(self.headers and self.table_data)
        self.reload_btn.setEnabled(
            not loading and has_file
        )
        self.open_excel_btn.setEnabled(not loading and has_file)
        self.sheet_combo.setEnabled(
            not loading and self.sheet_combo.count() > 0
        )
        self.name_column_combo.setEnabled(not loading and has_data)
        self.extract_columns_btn.setEnabled(not loading and has_data)
        self.wechat_column_combo.setEnabled(not loading and has_data)
        self.load_data_btn.setEnabled(not loading and has_data)
        self.save_config_btn.setEnabled(not loading and has_data)
        self.save_as_config_btn.setEnabled(not loading and has_data)
        self.load_config_btn.setEnabled(not loading)
        self.recent_config_list.setEnabled(not loading)

    def _load_excel_data(self, file_path, sheet_name=None):
        if self.excel_worker and self.excel_worker.isRunning():
            self.log("Excel文件正在读取，已忽略重复请求")
            return False

        self.log(f"选择Excel文件: {file_path}")
        self._set_excel_loading(True)

        worker = ExcelReadWorker(file_path, sheet_name)
        self.excel_worker = worker
        worker.signals.result.connect(
            lambda result, current=worker: self.on_excel_read_result(current, result)
        )
        worker.signals.error.connect(
            lambda error, current=worker: self.on_excel_read_error(current, error)
        )
        worker.signals.log.connect(self.log)
        worker.finished.connect(
            lambda current=worker: self.on_excel_read_finished(current)
        )
        worker.start()
        return True

    def on_excel_read_result(self, worker, result):
        if worker is not self.excel_worker:
            return

        self.sheet_names = result['sheet_names']
        self.current_sheet = result['current_sheet']
        self.headers = result['headers']
        self.table_data = result['data']
        self.current_excel_path = worker.file_path

        self.current_file_label.setText(
            f"当前文件: {os.path.basename(self.current_excel_path)}"
        )
        self.processor = None
        self.wechat_mapping = {}
        self.persons_list.clear()
        self.preview_text.clear()
        self.send_btn.setEnabled(False)
        self.start_send_btn.setEnabled(False)
        self.last_failed_tasks = []
        self.retry_send_btn.setEnabled(False)
        self.clear_filter_conditions()
        
        self.sheet_combo.blockSignals(True)
        try:
            self.sheet_combo.clear()
            self.sheet_combo.addItems(self.sheet_names)
            self.sheet_combo.setCurrentText(self.current_sheet)
        finally:
            self.sheet_combo.blockSignals(False)
        
        self._update_column_combos()
        if not self._apply_pending_config(worker):
            self.log("Excel文件读取完成！请选择Sheet和列，然后点击'加载数据'")
    
    def on_excel_read_error(self, worker, error):
        if worker is not self.excel_worker:
            return

        if self.current_sheet:
            self.sheet_combo.blockSignals(True)
            self.sheet_combo.setCurrentText(self.current_sheet)
            self.sheet_combo.blockSignals(False)
        QMessageBox.warning(self, "读取失败", error)
    
    def on_excel_read_finished(self, worker):
        if worker is self.excel_worker:
            self.excel_worker = None
            self._set_excel_loading(False)
        worker.deleteLater()
    
    def on_sheet_changed(self, index):
        if index >= 0 and hasattr(self, 'current_excel_path') and self.current_excel_path:
            sheet_name = self.sheet_combo.itemText(index)
            self.log(f"--- 切换到Sheet: {sheet_name} ---")
            
            if self.excel_worker and self.excel_worker.isRunning():
                self.log("Excel文件正在读取，请稍候再切换Sheet")
                return

            self._load_excel_data(self.current_excel_path, sheet_name)
    
    def _update_column_combos(self):
        self.name_column_combo.clear()
        self.name_column_combo.addItems(self.headers)
        
        self.wechat_column_combo.clear()
        self.wechat_column_combo.addItems([""] + self.headers)
        
        self.selected_columns = []
        self._update_selected_columns_label()
    
    def load_data(self):
        name_column = self.name_column_combo.currentText()
        extract_columns = self.selected_columns
        wechat_column = self.wechat_column_combo.currentText() if self.wechat_column_combo.currentIndex() > 0 else ""
        
        if not name_column:
            QMessageBox.warning(self, "警告", "请选择人名所在列")
            return
        
        if not extract_columns:
            QMessageBox.warning(self, "警告", "请至少选择一列要提取的数据")
            return
        
        self.processor = TableProcessor(self.table_data)
        self.refresh_persons_list()
        
        if wechat_column:
            try:
                self.wechat_mapping = self.processor.get_person_to_wechat_mapping(name_column, wechat_column)
                self.log(f"已建立 {len(self.wechat_mapping)} 个微信映射")
            except Exception as e:
                self.log(f"⚠ 建立微信映射失败: {e}")
        
        self.log(f"数据加载完成！找到 {self.persons_list.count()} 个人")

    def log(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())
        
        # 用 QTextDocument 原生上限做裁剪，CPU 比手动逐块删除更低
        if not hasattr(self, "_log_max_lines_applied"):
            self.log_text.document().setMaximumBlockCount(1200)
            self._log_max_lines_applied = True

    def send_data(self):
        selected_items = self.persons_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请选择要发送的人员")
            return
        
        extract_columns = self.selected_columns
        if not extract_columns:
            QMessageBox.warning(self, "警告", "请至少选择一列要提取的数据")
            return
        
        name_column = self.name_column_combo.currentText()
        
        custom_msg = ""
        if self.custom_msg_checkbox.isChecked():
            custom_msg = self.custom_msg_edit.toPlainText().strip()
        send_mode = self.send_mode_combo.currentData()
        
        tasks = []
        missing_wechat = []
        
        for item in selected_items:
            name = item.text()
            table_data = self.processor.get_person_table_data(
                name, 
                name_column, 
                extract_columns,
                self.filter_conditions
            )
            
            if not table_data:
                self.log(f"未找到 {name} 的数据，跳过")
                continue
            person_data = self.processor.format_table_data(table_data)
            
            recipient = self.wechat_mapping.get(name, "")
            if not recipient:
                recipient = self.wechat_edit.text().strip()
            
            if not recipient:
                recipient = name
            
            tasks.append({
                "name": name,
                "person_data": person_data,
                "table_data": table_data,
                "recipient": recipient,
                "custom_msg": custom_msg
            })
        
        if missing_wechat:
            QMessageBox.warning(self, "警告", f"以下人员缺少微信接收人，已跳过:\n{', '.join(missing_wechat)}")
        
        if not tasks:
            QMessageBox.warning(self, "警告", "没有可发送的任务")
            return
        
        max_display = 20
        confirm_text = f"即将向以下 {len(tasks)} 人发送消息:\n\n"
        for i, task in enumerate(tasks):
            if i >= max_display:
                confirm_text += f"  • ... 还有 {len(tasks) - max_display} 人\n"
                break
            confirm_text += f"  • {task['name']} → {task['recipient']}\n"
        confirm_text += f"\n发送间隔: {self.send_interval_spin.value()}秒"
        confirm_text += f"\n数据发送形式: {self.send_mode_combo.currentText()}"
        if custom_msg:
            confirm_text += f"\n包含自定义消息: {custom_msg[:30]}..." if len(custom_msg) > 30 else f"\n包含自定义消息: {custom_msg}"
        
        reply = QMessageBox.question(
            self, "确认发送", confirm_text,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply != QMessageBox.StandardButton.Yes:
            self.log("用户取消发送")
            return
        
        self.send_btn.setEnabled(False)
        self.start_send_btn.setEnabled(False)
        self.pause_send_btn.setEnabled(True)
        self.stop_send_btn.setEnabled(True)
        self.progress_bar.setValue(0)
        self.progress_label.setText("正在发送...")
        self.last_send_mode = send_mode
        
        worker = SendWorker(
            tasks=tasks,
            send_interval=self.send_interval_spin.value(),
            chat_delay=self.chat_delay_spin.value(),
            send_mode=send_mode,
            minimize_after=self.minimize_check.isChecked(),
            attachment=self.attachment_edit.text().strip()
        )
        self.worker = worker
        worker.signals.result.connect(self.on_send_result)
        worker.signals.error.connect(self.on_send_error)
        worker.signals.log.connect(self.log)
        worker.signals.progress.connect(self.on_send_progress)
        worker.finished.connect(
            lambda current=worker: self.on_send_finished(current)
        )
        worker.start()

    def pause_send(self):
        if self.worker and self.worker.isRunning():
            if self.worker.is_paused():
                self.worker.set_paused(False)
                self.pause_send_btn.setText("⏸ 暂停")
                self.log("▶ 继续发送")
            else:
                self.worker.set_paused(True)
                self.pause_send_btn.setText("▶ 继续")
                self.log("⏸ 发送已暂停")
    
    def stop_send(self):
        if self.worker and self.worker.isRunning():
            self.worker.set_stopped(True)
            self.worker.set_paused(False)
            self.log("⏹ 正在停止发送，当前操作结束后将清理临时图片...")

    def retry_send(self):
        if self.worker and self.worker.isRunning():
            self.log("发送线程仍在结束，请稍候再重试")
            return

        if not hasattr(self, 'last_failed_tasks') or not self.last_failed_tasks:
            QMessageBox.warning(self, "警告", "没有可重试的任务")
            return
        
        failed_tasks = self.last_failed_tasks
        self.log(f"🔄 开始重试发送，共 {len(failed_tasks)} 个失败任务")
        
        self.send_btn.setEnabled(False)
        self.start_send_btn.setEnabled(False)
        self.retry_send_btn.setEnabled(False)
        self.pause_send_btn.setEnabled(True)
        self.stop_send_btn.setEnabled(True)
        self.progress_bar.setValue(0)
        self.progress_label.setText("正在重试发送...")
        
        worker = SendWorker(
            tasks=failed_tasks,
            send_interval=self.send_interval_spin.value(),
            chat_delay=self.chat_delay_spin.value(),
            send_mode=getattr(self, 'last_send_mode', 'text'),
            minimize_after=self.minimize_check.isChecked(),
            attachment=self.attachment_edit.text().strip()
        )
        self.worker = worker
        worker.signals.result.connect(self.on_send_result)
        worker.signals.error.connect(self.on_send_error)
        worker.signals.log.connect(self.log)
        worker.signals.progress.connect(self.on_send_progress)
        worker.finished.connect(
            lambda current=worker: self.on_send_finished(current)
        )
        worker.start()

    def on_send_finished(self, worker):
        if worker is self.worker:
            self.worker = None
            self.send_btn.setEnabled(True)
            self.start_send_btn.setEnabled(True)
            self.pause_send_btn.setEnabled(False)
            self.stop_send_btn.setEnabled(False)
            self.pause_send_btn.setText("⏸ 暂停")
        worker.deleteLater()

    def on_send_progress(self, value):
        self.progress_bar.setValue(value)
        self.progress_label.setText(f"发送进度: {value}%")

    def on_send_result(self, result):
        success_count, failed_count, total_count, failed_tasks = result
        self.last_failed_tasks = failed_tasks
        QMessageBox.information(self, "完成", f"发送完成！成功 {success_count}, 失败 {failed_count}, 总计 {total_count}")
        self.progress_label.setText(f"发送完成: 成功 {success_count}, 失败 {failed_count}")
        
        if failed_tasks:
            self.retry_send_btn.setEnabled(True)
            self.log(f"⚠ 有 {len(failed_tasks)} 个发送失败，可点击'重试发送'按钮重新发送")
        else:
            self.retry_send_btn.setEnabled(False)

    def on_send_error(self, error):
        QMessageBox.critical(self, "错误", f"发送失败: {error}")
        self.log(f"发送失败: {error}")
        self.progress_label.setText("发送出错")


class MultiSelectDialog(QDialog):
    def __init__(self, title, items, selected_items=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setGeometry(300, 300, 400, 350)
        
        self.selected_columns = []
        self.init_ui(items, selected_items or [])
    
    def init_ui(self, items, selected_items):
        layout = QVBoxLayout()
        
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        
        for item in items:
            list_item = QListWidgetItem(item)
            if item in selected_items:
                list_item.setSelected(True)
            self.list_widget.addItem(list_item)
        
        layout.addWidget(self.list_widget)
        
        btn_layout = QHBoxLayout()
        
        select_all_btn = QPushButton("全选")
        select_all_btn.clicked.connect(self.select_all)
        btn_layout.addWidget(select_all_btn)
        
        deselect_all_btn = QPushButton("取消全选")
        deselect_all_btn.clicked.connect(self.deselect_all)
        btn_layout.addWidget(deselect_all_btn)
        
        layout.addLayout(btn_layout)
        
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def select_all(self):
        self.list_widget.selectAll()
    
    def deselect_all(self):
        self.list_widget.clearSelection()
    
    def get_selected(self):
        return [item.text() for item in self.list_widget.selectedItems()]
    
    def accept(self):
        self.selected_columns = self.get_selected()
        super().accept()


class ScheduleTab(QWidget):
    """定时发送标签页：任务列表、编辑表单、保存/加载配置、手动立即发送。"""

    # worker 线程日志 → 信号投递回主线程再写入日志面板（QThread 无 event loop，
    # QTimer.singleShot 不会触发；signal.emit 走 QueuedConnection 自动排队到主线程）
    _worker_log_signal = pyqtSignal(str)

    def __init__(self, store: ScheduleStore, dispatcher: ScheduleDispatcher, parent=None):
        super().__init__(parent)
        self.store = store
        self.dispatcher = dispatcher
        self.tasks: Dict[str, ScheduleTask] = {}
        self.current_task_id: Optional[str] = None
        self._worker_log_signal.connect(self._on_worker_log)
        self.send_worker: Optional[ScheduleSendWorker] = None
        # 接收 MainWindow 的统一日志回调，外部赋值
        self.log_callback = None
        self.init_ui()
        self.connect_signals()
        self.reload_tasks()

    # ----------------------------- UI -----------------------------
    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(8)

        # --- 左侧：任务列表 + 通用按钮 ---
        left_group = QGroupBox("定时任务列表")
        left_layout = QVBoxLayout(left_group)

        self.task_list = QListWidget()
        self.task_list.setMinimumWidth(240)
        left_layout.addWidget(self.task_list)

        left_btn1 = QHBoxLayout()
        self.add_task_btn = QPushButton("新建任务")
        self.clone_task_btn = QPushButton("复制任务")
        self.delete_task_btn = QPushButton("删除任务")
        left_btn1.addWidget(self.add_task_btn)
        left_btn1.addWidget(self.clone_task_btn)
        left_btn1.addWidget(self.delete_task_btn)
        left_layout.addLayout(left_btn1)

        left_btn2 = QHBoxLayout()
        self.save_all_btn = QPushButton("保存为定时配置")
        self.load_config_btn = QPushButton("加载定时配置")
        left_btn2.addWidget(self.save_all_btn)
        left_btn2.addWidget(self.load_config_btn)
        left_layout.addLayout(left_btn2)

        left_btn3 = QHBoxLayout()
        self.run_now_btn = QPushButton("立即执行所选任务")
        self.stop_send_btn = QPushButton("停止当前发送")
        self.stop_send_btn.setEnabled(False)
        left_btn3.addWidget(self.run_now_btn)
        left_btn3.addWidget(self.stop_send_btn)
        left_layout.addLayout(left_btn3)

        left_btn4 = QHBoxLayout()
        self.weather_settings_btn = QPushButton("🌤 天气设置")
        left_btn4.addWidget(self.weather_settings_btn)
        left_btn4.addStretch()
        left_layout.addLayout(left_btn4)

        self.dispatcher_status_label = QLabel("调度器状态: 未启动")
        self.dispatcher_status_label.setStyleSheet("color:#666; font-size:11px;")
        left_layout.addWidget(self.dispatcher_status_label)

        left_layout.addStretch()
        main_layout.addWidget(left_group, 0)

        # --- 右侧：编辑表单 ---
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        # 基本
        base_group = QGroupBox("任务基础")
        base_layout = QVBoxLayout(base_group)
        row = QHBoxLayout()
        row.addWidget(QLabel("任务名称:"))
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("例如：周一早会提醒")
        row.addWidget(self.name_edit, 1)
        self.enabled_check = QCheckBox("启用")
        self.enabled_check.setChecked(True)
        row.addWidget(self.enabled_check)
        base_layout.addLayout(row)
        right_layout.addWidget(base_group)

        # 重复规则
        repeat_group = QGroupBox("重复规则")
        repeat_layout = QVBoxLayout(repeat_group)
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("模式:"))
        self.repeat_mode_combo = QComboBox()
        self.repeat_mode_combo.addItem("每天", "daily")
        self.repeat_mode_combo.addItem("每周指定日期", "weekly")
        row1.addWidget(self.repeat_mode_combo, 1)
        repeat_layout.addLayout(row1)

        self.weekday_group_box = QGroupBox("选择周几(每周模式生效):")
        weekday_layout = QHBoxLayout(self.weekday_group_box)
        self.weekday_checks: Dict[int, QCheckBox] = {}
        for idx, name in enumerate(WEEKDAY_NAMES, start=1):
            cb = QCheckBox(name)
            self.weekday_checks[idx] = cb
            weekday_layout.addWidget(cb)
        repeat_layout.addWidget(self.weekday_group_box)
        self._on_repeat_mode_changed(self.repeat_mode_combo.currentIndex())

        time_group = QGroupBox("发送时间点(HH:MM)，点击右侧按钮新增/删除:")
        time_layout = QVBoxLayout(time_group)
        self.time_list = QListWidget()
        self.time_list.setMaximumHeight(72)
        time_layout.addWidget(self.time_list)
        time_row = QHBoxLayout()
        self.time_edit = QTimeEdit()
        self.time_edit.setDisplayFormat("HH:mm")
        self.time_edit.setTime(QTime(9, 0))
        time_row.addWidget(self.time_edit)
        self.add_time_btn = QPushButton("添加")
        self.del_time_btn = QPushButton("删除选中")
        time_row.addWidget(self.add_time_btn)
        time_row.addWidget(self.del_time_btn)
        time_layout.addLayout(time_row)
        repeat_layout.addWidget(time_group)
        right_layout.addWidget(repeat_group)

        # 发送内容
        send_group = QGroupBox("发送内容")
        send_layout = QVBoxLayout(send_group)

        send_layout.addWidget(QLabel("接收人(好友/群名，支持模糊匹配，每行一个或英文逗号分隔):"))
        self.recipients_edit = QTextEdit()
        self.recipients_edit.setMaximumHeight(70)
        self.recipients_edit.setPlaceholderText("张三\n文件传输助手\n工作群A")
        send_layout.addWidget(self.recipients_edit)

        send_layout.addWidget(QLabel("自定义文字消息:"))
        self.message_edit = QTextEdit()
        self.message_edit.setMinimumHeight(110)
        self.message_edit.setPlaceholderText(
            "早安！今天记得填写日报。\n"
            "支持占位符：\n"
            "  天气(需在「天气设置」配置 API key)：\n"
            "    {{weather}}        简洁：多云 30°C\n"
            "    {{weather_full}}   完整：实时+最高最低+白天夜间+风力+建议\n"
            "    {{weather:北京}}   指定城市\n"
            "    {{weather_max}} {{weather_min}}  最高/最低温\n"
            "    {{weather_day}} {{weather_night}} 白天/夜间天气\n"
            "    {{weather_wind}} {{weather_advice}} 风力/生活建议\n"
            "  热搜(无 key)：{{news}} {{news:zhihu}} {{news:toutiao}} {{news:bilibili}} {{news:5}}\n"
            "  时间：{{date}} {{weekday}} {{time}}"
        )
        send_layout.addWidget(self.message_edit)

        city_row = QHBoxLayout()
        city_row.addWidget(QLabel("默认城市(天气占位符不带参时用此):"))
        self.default_city_edit = QLineEdit()
        self.default_city_edit.setPlaceholderText("如：北京；留空则用「天气设置」里的全局默认城市")
        city_row.addWidget(self.default_city_edit, 1)
        send_layout.addLayout(city_row)

        attach_row = QHBoxLayout()
        attach_row.addWidget(QLabel("附加文件:"))
        self.attachment_edit = QLineEdit()
        self.attachment_edit.setReadOnly(True)
        self.attachment_edit.setPlaceholderText("可选，选择文件/图片随消息一起发送")
        attach_row.addWidget(self.attachment_edit, 1)
        self.attach_pick_btn = QPushButton("选择...")
        self.attach_clear_btn = QPushButton("清除")
        self.attach_pick_btn.setFixedWidth(60)
        self.attach_clear_btn.setFixedWidth(50)
        attach_row.addWidget(self.attach_pick_btn)
        attach_row.addWidget(self.attach_clear_btn)
        send_layout.addLayout(attach_row)

        delay_row = QHBoxLayout()
        delay_row.addWidget(QLabel("聊天窗口切换延迟(秒):"))
        self.chat_delay_spin = QDoubleSpinBox()
        self.chat_delay_spin.setRange(0.0, 10.0)
        self.chat_delay_spin.setSingleStep(0.1)
        self.chat_delay_spin.setValue(0.3)
        delay_row.addWidget(self.chat_delay_spin)
        delay_row.addWidget(QLabel("每人发送间隔(秒):"))
        self.send_interval_spin = QDoubleSpinBox()
        self.send_interval_spin.setRange(0.0, 30.0)
        self.send_interval_spin.setSingleStep(0.1)
        self.send_interval_spin.setValue(0.5)
        delay_row.addWidget(self.send_interval_spin)
        delay_row.addStretch()
        send_layout.addLayout(delay_row)

        right_layout.addWidget(send_group)

        # 电脑锁定/完成后行为配置（任务级，每个任务可单独设置）
        lock_group = QGroupBox("电脑锁定 / 完成后行为")
        lock_layout = QHBoxLayout(lock_group)
        self.keep_unlock_check = QCheckBox("执行期间防自动锁定")
        self.keep_unlock_check.setToolTip(
            "开启后：任务执行期间阻止电脑自动锁定/休眠，确保发送不被打断。\n"
            "注意：若电脑已被手动锁定(Win+L)，程序无法自动解锁，任务会失败。"
        )
        self.relock_check = QCheckBox("发送完成后自动锁定")
        self.relock_check.setToolTip(
            "开启后：任务发送完成后自动锁定电脑（等同 Win+L）。\n"
            "若近期 15 分钟内还有其他定时任务，会等最后一个任务完成后再锁定。"
        )
        self.minimize_check = QCheckBox("发送后最小化微信")
        self.minimize_check.setToolTip(
            "开启后：整个任务发送完成后把微信窗口最小化一次（保护隐私）。\n"
            "关闭后微信窗口保持原样，不会被最小化。"
        )
        self.minimize_check.setChecked(True)
        lock_layout.addWidget(self.keep_unlock_check)
        lock_layout.addWidget(self.relock_check)
        lock_layout.addWidget(self.minimize_check)
        lock_layout.addStretch()
        right_layout.addWidget(lock_group)

        # 操作区 + 保存
        action_row = QHBoxLayout()
        self.save_task_btn = QPushButton("💾 保存当前任务")
        self.save_task_btn.setStyleSheet(
            "background-color:#2196F3; color:white; padding:6px;"
        )
        self.reset_form_btn = QPushButton("重置表单")
        action_row.addWidget(self.save_task_btn)
        action_row.addWidget(self.reset_form_btn)
        action_row.addStretch()
        right_layout.addLayout(action_row)

        # 日志（共享主日志的回调，这里也放只读面板，方便查看）
        log_group = QGroupBox("定时发送日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(140)
        log_layout.addWidget(self.log_text)
        right_layout.addWidget(log_group, 1)

        main_layout.addWidget(right, 1)

    # ------------------------- 信号绑定 -------------------------
    def connect_signals(self):
        self.task_list.currentItemChanged.connect(self._on_task_selected)
        self.task_list.itemChanged.connect(self._on_task_item_changed)
        self.add_task_btn.clicked.connect(self._on_new_task)
        self.clone_task_btn.clicked.connect(self._on_clone_task)
        self.delete_task_btn.clicked.connect(self._on_delete_task)
        self.save_task_btn.clicked.connect(self._on_save_current_task)
        self.reset_form_btn.clicked.connect(self._reset_form)
        self.add_time_btn.clicked.connect(self._on_add_time)
        self.del_time_btn.clicked.connect(self._on_del_time)
        self.repeat_mode_combo.currentIndexChanged.connect(self._on_repeat_mode_changed)
        self.attach_pick_btn.clicked.connect(self._on_pick_attachment)
        self.attach_clear_btn.clicked.connect(self._on_clear_attachment)
        self.save_all_btn.clicked.connect(self._on_save_all)
        self.load_config_btn.clicked.connect(self._on_load_all)
        self.run_now_btn.clicked.connect(self._on_run_now)
        self.stop_send_btn.clicked.connect(self._on_stop_send)
        self.weather_settings_btn.clicked.connect(self._on_open_weather_settings)

    # ------------------------- 日志 -------------------------
    # 最大日志行数（交给 QTextDocument 原生裁剪，CPU/内存都更省）
    _MAX_LOG_LINES = 1200

    def _on_worker_log(self, message: str) -> None:
        # 主线程槽：接收 worker 线程通过 _worker_log_signal 投递的日志
        self.log(message)

    def log(self, message: str) -> None:
        # 本地追加
        if hasattr(self, "log_text") and self.log_text is not None:
            from datetime import datetime
            ts = datetime.now().strftime("%H:%M:%S")
            self.log_text.append(f"[{ts}] {message}")
            # 只设置一次即可，后续由 QTextDocument 原生自动裁剪
            if not getattr(self, "_log_max_applied", False):
                self.log_text.document().setMaximumBlockCount(self._MAX_LOG_LINES)
                self._log_max_applied = True
        # 同时推送给 MainWindow，用于主界面统一日志面板
        if self.log_callback:
            try:
                self.log_callback(message)
            except Exception:
                pass

    # ------------------------- 任务列表 -------------------------
    def reload_tasks(self):
        self.task_list.blockSignals(True)
        self.task_list.clear()
        tasks = self.store.load_all()
        self.tasks = {t.id: t for t in tasks}
        for t in tasks:
            item = QListWidgetItem(self._format_task_label(t))
            item.setData(Qt.ItemDataRole.UserRole, t.id)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(
                Qt.CheckState.Checked if t.enabled else Qt.CheckState.Unchecked
            )
            self.task_list.addItem(item)
        self.task_list.blockSignals(False)
        if tasks:
            self.task_list.setCurrentRow(0)
        else:
            self._reset_form()

    def _format_task_label(self, task: ScheduleTask) -> str:
        mark = "●" if task.enabled else "○"
        if task.repeat_mode == "daily":
            rule = "每天"
        else:
            days = ",".join(WEEKDAY_NAMES[d - 1] for d in task.days) if task.days else "未选"
            rule = f"每周: {days}"
        times = "、".join(task.times) if task.times else "(无时间点)"
        count = len(task.recipients)
        return f"{mark} {task.name}  | {rule} | {times} | 人数:{count}"

    def _current_task_id_from_list(self) -> Optional[str]:
        item = self.task_list.currentItem()
        if item is None:
            return None
        return item.data(Qt.ItemDataRole.UserRole)

    def _on_task_selected(self, current, previous):
        tid = None
        if current:
            tid = current.data(Qt.ItemDataRole.UserRole)
        if not tid or tid not in self.tasks:
            self.current_task_id = None
            return
        self.current_task_id = tid
        self._load_task_to_form(self.tasks[tid])

    def _on_task_item_changed(self, item):
        tid = item.data(Qt.ItemDataRole.UserRole)
        if not tid:
            return
        task = self.tasks.get(tid)
        if not task:
            return
        new_enabled = item.checkState() == Qt.CheckState.Checked
        if task.enabled == new_enabled:
            return
        task.enabled = new_enabled
        self.store.save(task)
        self.tasks[tid] = task
        self.task_list.blockSignals(True)
        try:
            item.setText(self._format_task_label(task))
        finally:
            self.task_list.blockSignals(False)
        self.log(f"[定时] {task.name} 已{'启用' if new_enabled else '禁用'}")

    def _load_task_to_form(self, task: ScheduleTask):
        self.name_edit.setText(task.name)
        self.enabled_check.setChecked(task.enabled)
        idx = 0 if task.repeat_mode == "daily" else 1
        self.repeat_mode_combo.setCurrentIndex(idx)
        for d, cb in self.weekday_checks.items():
            cb.setChecked(d in set(task.days))
        self.time_list.clear()
        for slot in task.times:
            self.time_list.addItem(slot)
        self.recipients_edit.setPlainText("\n".join(task.recipients))
        self.message_edit.setPlainText(task.message)
        self.chat_delay_spin.setValue(task.chat_delay)
        self.send_interval_spin.setValue(task.send_interval)
        self.default_city_edit.setText(task.default_city or "")
        self.keep_unlock_check.setChecked(bool(getattr(task, "keep_unlocked", False)))
        self.relock_check.setChecked(bool(getattr(task, "relock_after", False)))
        self.minimize_check.setChecked(bool(getattr(task, "minimize_after", True)))
        self.attachment_edit.setText(getattr(task, "attachment", "") or "")
        self._on_repeat_mode_changed(idx)

    def _reset_form(self):
        self.current_task_id = None
        self.name_edit.clear()
        self.enabled_check.setChecked(True)
        self.repeat_mode_combo.setCurrentIndex(0)
        for cb in self.weekday_checks.values():
            cb.setChecked(False)
        self.time_list.clear()
        self.time_edit.setTime(QTime(9, 0))
        self.recipients_edit.clear()
        self.message_edit.clear()
        self.chat_delay_spin.setValue(0.3)
        self.send_interval_spin.setValue(0.5)
        self.default_city_edit.clear()
        self.keep_unlock_check.setChecked(False)
        self.relock_check.setChecked(False)
        self.minimize_check.setChecked(True)
        self.attachment_edit.clear()

    # ------------------------- 增删改 -------------------------
    def _on_new_task(self):
        self._reset_form()
        self.name_edit.setFocus()

    def _on_clone_task(self):
        task = self._form_to_task(new_id=True)
        task.name = f"{task.name} - 副本"
        self.store.save(task)
        self.reload_tasks()
        self._select_task_by_id(task.id)
        self.log(f"[定时] 已复制任务: {task.name}")

    def _on_delete_task(self):
        tid = self._current_task_id_from_list()
        if not tid:
            QMessageBox.information(self, "提示", "请先选择要删除的任务")
            return
        task = self.tasks.get(tid)
        name = task.name if task else tid
        ans = QMessageBox.question(
            self,
            "删除确认",
            f"确定要删除任务「{name}」吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if ans != QMessageBox.StandardButton.Yes:
            return
        if self.store.delete(tid):
            self.log(f"[定时] 已删除任务: {name}")
            self.reload_tasks()

    def _on_save_current_task(self):
        try:
            task = self._form_to_task(new_id=self.current_task_id is None)
        except ValueError as exc:
            QMessageBox.warning(self, "任务信息不完整", str(exc))
            return
        if self.current_task_id:
            task.id = self.current_task_id
        self.store.save(task)
        self.current_task_id = task.id
        self.log(f"[定时] 已保存任务: {task.name}")
        self.reload_tasks()
        self._select_task_by_id(task.id)

    def _select_task_by_id(self, task_id):
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == task_id:
                self.task_list.setCurrentRow(i)
                return

    def _form_to_task(self, new_id: bool = True) -> ScheduleTask:
        name = self.name_edit.text().strip() or "未命名任务"
        repeat_mode = self.repeat_mode_combo.currentData()
        days: List[int] = []
        if repeat_mode == "weekly":
            days = sorted([d for d, cb in self.weekday_checks.items() if cb.isChecked()])
            if not days:
                raise ValueError("每周模式请至少选择一个周几")
        times: List[str] = []
        for i in range(self.time_list.count()):
            slot = _normalize_time(self.time_list.item(i).text())
            if slot and slot not in times:
                times.append(slot)
        if not times:
            raise ValueError("请至少添加一个发送时间点")
        recipients_raw = self.recipients_edit.toPlainText()
        recipients = _parse_recipients(recipients_raw)
        if not recipients:
            raise ValueError("请至少填写一个接收人(好友名/群名)")
        message = self.message_edit.toPlainText()
        if not message.strip():
            raise ValueError("请填写自定义文字消息")
        task = ScheduleTask(
            id=_new_task_id() if new_id else (self.current_task_id or _new_task_id()),
            name=name,
            enabled=self.enabled_check.isChecked(),
            repeat_mode=repeat_mode,
            days=days,
            times=times,
            recipients=recipients,
            message=message,
            chat_delay=float(self.chat_delay_spin.value()),
            send_interval=float(self.send_interval_spin.value()),
            default_city=self.default_city_edit.text().strip(),
            keep_unlocked=self.keep_unlock_check.isChecked(),
            relock_after=self.relock_check.isChecked(),
            minimize_after=self.minimize_check.isChecked(),
            attachment=self.attachment_edit.text().strip(),
        )
        return task

    # ------------------------- 附加文件 -------------------------
    def _on_pick_attachment(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择附加文件(文件/图片)",
            "",
            "常用文件 (*.png *.jpg *.jpeg *.gif *.bmp *.webp *.pdf *.doc *.docx *.xls *.xlsx *.ppt *.pptx *.txt *.zip *.rar *.7z);;所有文件 (*.*)",
        )
        if path:
            self.attachment_edit.setText(path)

    def _on_clear_attachment(self):
        self.attachment_edit.clear()

    # ------------------------- 重复模式/时间点 -------------------------
    def _on_repeat_mode_changed(self, index):
        weekly = self.repeat_mode_combo.currentData() == "weekly"
        self.weekday_group_box.setEnabled(weekly)

    def _on_add_time(self):
        t = self.time_edit.time().toString("HH:mm")
        slot = _normalize_time(t)
        if not slot:
            return
        for i in range(self.time_list.count()):
            if self.time_list.item(i).text() == slot:
                return
        self.time_list.addItem(slot)

    def _on_del_time(self):
        for item in list(self.time_list.selectedItems()):
            self.time_list.takeItem(self.time_list.row(item))

    # ------------------------- 配置保存/加载（独立 JSON 文件）-------------------------
    def _collect_all_tasks(self):
        # 先把当前编辑中的任务表单保存到内存中再导出？不，只导出已保存的列表
        return self.store.load_all()

    def _on_save_all(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "保存定时发送配置",
            os.path.join(os.path.expanduser("~"), "定时发送配置.json"),
            "JSON 配置 (*.json)",
        )
        if not path:
            return
        tasks = self._collect_all_tasks()
        data = {
            "version": 1,
            "type": "schedule_profile",
            "saved_at": "",
            "tasks": [t.to_dict() for t in tasks],
        }
        try:
            from modules.schedule_manager import _write_json
            from datetime import datetime
            data["saved_at"] = datetime.now().isoformat(timespec="seconds")
            _write_json(path, data)
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", str(exc))
            return
        self.log(f"[定时] 配置已保存到: {path}")
        QMessageBox.information(self, "保存成功", f"已保存 {len(tasks)} 个定时任务配置")

    def _on_load_all(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "加载定时发送配置", "", "JSON 配置 (*.json)"
        )
        if not path:
            return
        try:
            import json
            with open(path, "r", encoding="utf-8") as fp:
                raw = json.load(fp)
            items = raw.get("tasks") if isinstance(raw, dict) else None
            if not isinstance(items, list):
                raise ValueError("配置文件中没有 tasks 列表")
            loaded = [ScheduleTask.from_dict(it) for it in items]
        except Exception as exc:
            QMessageBox.critical(self, "加载失败", f"读取配置失败: {exc}")
            return
        ans = QMessageBox.question(
            self,
            "导入方式",
            "是否覆盖现有定时任务？\n选择【是】=覆盖；【否】=追加。",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No
            | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel,
        )
        if ans == QMessageBox.StandardButton.Cancel:
            return
        overwrite = ans == QMessageBox.StandardButton.Yes
        if overwrite:
            for old in self.store.load_all():
                self.store.delete(old.id)
        for t in loaded:
            # 避免 ID 冲突时覆盖，统一生成新 ID
            t.id = _new_task_id()
            self.store.save(t)
        self.reload_tasks()
        self.log(f"[定时] 已导入 {len(loaded)} 个任务配置: {os.path.basename(path)}")
        QMessageBox.information(self, "导入成功", f"已导入 {len(loaded)} 个定时任务")

    # ------------------------- 立即执行 -------------------------
    def _on_run_now(self):
        if self.send_worker and self.send_worker.isRunning():
            QMessageBox.information(self, "提示", "已有定时发送正在进行中")
            return
        tid = self._current_task_id_from_list()
        if not tid:
            QMessageBox.information(self, "提示", "请先选择要执行的任务")
            return
        task = self.tasks.get(tid)
        if not task:
            return
        # 保存当前表单到任务（保留用户临时修改）
        try:
            current = self._form_to_task(new_id=False)
            current.id = task.id
        except ValueError as exc:
            QMessageBox.warning(self, "任务信息不完整", str(exc))
            return

        # 关键：ScheduleSendWorker 在子线程里调用 log_callback，不能直接碰 QWidget。
        # 这里把日志投递切回主线程，避免跨线程访问控件导致的偶发崩溃/挂起。
        worker = ScheduleSendWorker(
            recipients=current.recipients,
            message=current.message,
            chat_delay=current.chat_delay,
            send_interval=current.send_interval,
            log_callback=self._worker_log_signal.emit,
            default_city=current.default_city,
            minimize_after=getattr(current, "minimize_after", True),
            attachment=getattr(current, "attachment", "") or "",
        )

        # 保存任务配置供完成回调使用
        self._current_run_task = current

        worker.finished_with_result.connect(self._on_send_finished)
        # 让 QThread 自动回收 C++ 对象，避免反复启动后 Qt 对象堆积
        worker.finished.connect(lambda: self._cleanup_schedule_worker_after(worker))
        self.send_worker = worker
        self.stop_send_btn.setEnabled(True)
        self.run_now_btn.setEnabled(False)

        # 按任务配置启动防锁定守护
        if getattr(current, "keep_unlocked", False):
            mw = self.window()
            if mw and hasattr(mw, "workstation_guard"):
                if not mw.workstation_guard.is_running():
                    mw.workstation_guard.start()
                    self.log(f"[定时] 任务「{current.name}」已开启防锁定守护")

        self.log(f"[定时] 手动立即执行任务: {current.name}")
        worker.start()

    def _cleanup_schedule_worker_after(self, worker):
        try:
            if worker is self.send_worker:
                # finished 之后已经 isRunning() == False，可以安全解除引用
                self.send_worker = None
        finally:
            try:
                worker.deleteLater()
            except Exception:
                pass

    def _on_stop_send(self):
        if self.send_worker and self.send_worker.isRunning():
            self.send_worker.stop()
            self.log("[定时] 已请求停止当前发送")

    def _on_send_finished(self, success: int, failed: int, failed_recipients: list):
        self.stop_send_btn.setEnabled(False)
        self.run_now_btn.setEnabled(True)
        msg = f"[定时] 本轮完成，成功: {success}，失败: {failed}"
        if failed_recipients:
            msg += f"（{'、'.join(str(r) for r in failed_recipients[:5])}"
            if len(failed_recipients) > 5:
                msg += f"…等{len(failed_recipients)}人"
            msg += "）"
        self.log(msg)
        # 按任务配置决定是否发送后锁定
        task = getattr(self, "_current_run_task", None)
        if task and getattr(task, "relock_after", False):
            mw = self.window()
            if mw and hasattr(mw, "workstation_guard"):
                try:
                    mw.workstation_guard.maybe_relock_after_task(mw.schedule_store)
                except Exception:
                    pass

    # ------------------------- 天气设置 -------------------------
    def _on_open_weather_settings(self):
        from modules import weather_fetcher
        dlg = QDialog(self)
        dlg.setWindowTitle("天气设置 - 和风天气 API")
        dlg.setMinimumWidth(480)
        layout = QVBoxLayout(dlg)

        cfg = weather_fetcher.load_weather_config()

        layout.addWidget(QLabel("和风天气 API Key:"))
        key_edit = QLineEdit(cfg.get("api_key", ""))
        key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        key_edit.setPlaceholderText("在和风天气控制台 https://console.qweather.com 获取")
        layout.addWidget(key_edit)

        layout.addWidget(QLabel("接入域名(免费版 devapi.qweather.com，商业版 api.qweather.com):"))
        url_edit = QLineEdit(cfg.get("base_url", weather_fetcher.DEFAULT_BASE_URL))
        url_edit.setPlaceholderText(weather_fetcher.DEFAULT_BASE_URL)
        layout.addWidget(url_edit)

        layout.addWidget(QLabel("全局默认城市(任务表单里没填默认城市时用此):"))
        city_edit = QLineEdit(cfg.get("default_city", ""))
        city_edit.setPlaceholderText("如：北京 / 上海 / 杭州")
        layout.addWidget(city_edit)

        result_label = QLabel("")
        result_label.setWordWrap(True)
        layout.addWidget(result_label)

        btn_row = QHBoxLayout()
        test_btn = QPushButton("测试连接")
        save_btn = QPushButton("💾 保存")
        cancel_btn = QPushButton("取消")
        btn_row.addWidget(test_btn)
        btn_row.addStretch()
        btn_row.addWidget(save_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

        def on_test():
            test_city = city_edit.text().strip() or "北京"
            self.log(f"[天气] 开始测试连接：city={test_city} url={url_edit.text().strip()}")
            try:
                ok, m = weather_fetcher.test_connection(
                    key_edit.text().strip(),
                    url_edit.text().strip(),
                    test_city,
                    log_fn=lambda msg: self.log(f"[天气] {msg}"),
                )
            except Exception as exc:
                ok, m = False, f"测试异常: {exc}"
                self.log(f"[天气] ❌ {m}")
            color = "#2e7d32" if ok else "#c62828"
            result_label.setText(f'<span style="color:{color}">{m}</span>')
            result_label.setTextFormat(Qt.TextFormat.RichText)

        def on_save():
            try:
                weather_fetcher.save_weather_config(
                    api_key=key_edit.text().strip(),
                    base_url=url_edit.text().strip(),
                    default_city=city_edit.text().strip(),
                )
                self.log("[定时] 天气配置已保存")
                dlg.accept()
            except Exception as exc:
                QMessageBox.critical(self, "保存失败", str(exc))

        test_btn.clicked.connect(on_test)
        save_btn.clicked.connect(on_save)
        cancel_btn.clicked.connect(dlg.reject)
        dlg.exec()


def _parse_recipients(text: str) -> List[str]:
    recipients: List[str] = []
    if not text:
        return recipients
    raw = str(text).replace("，", ",").replace(";", ",").replace("；", ",")
    for line in raw.splitlines():
        for part in line.split(","):
            name = part.strip()
            if name and name not in recipients:
                recipients.append(name)
    return recipients


class MainWindow(QMainWindow):
    # 跨线程日志投递：子线程 emit -> 主线程 slot 写 UI
    _schedule_log_signal = pyqtSignal(str)
    # 跨线程触发投递：调度器线程 emit -> 主线程创建 QThread（确保 affinity 正确）
    _schedule_trigger_signal = pyqtSignal(object, str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setGeometry(50, 50, 850, 580)
        self._closing_requested = False
        self._close_ready = False
        self._shutdown_poll_count = 0
        self._tray_hint_shown = False

        self.schedule_store = ScheduleStore()
        self.schedule_dispatcher = ScheduleDispatcher(self.schedule_store)
        self.schedule_dispatcher.add_log_handler(self._schedule_dispatch_log)
        self.schedule_dispatcher.add_handler(self._on_schedule_triggered)
        # 定时触发时并发串行化（避免同时多个任务抢微信窗口）
        self._schedule_running_lock = threading.Lock()
        self._schedule_worker = None

        # 关键：跨线程日志用 pyqtSignal 投递，不能用 QTimer.singleShot(0,...)
        # 因为子线程没有 Qt event loop，singleShot 永远不会触发
        self._schedule_log_signal.connect(self._on_schedule_log_signal)
        self._schedule_trigger_signal.connect(self._on_schedule_trigger_in_main)

        # 电脑锁定守护：阻止定时发送期间电脑自动锁定，任务完成后自动锁回
        self.workstation_guard = WorkstationGuard(
            log_fn=self._post_schedule_log_from_worker
        )

        self.init_ui()
        self.init_tray()
        self.schedule_dispatcher.start()
        self.schedule_tab.dispatcher_status_label.setText("调度器状态: 运行中")

    def init_ui(self):
        self.table_filter_tab = TableFilterTab()
        self.schedule_tab = ScheduleTab(self.schedule_store, self.schedule_dispatcher)
        # 让 ScheduleTab 日志同时写入原表格发送主界面的日志面板，保持统一查看
        self.schedule_tab.log_callback = self._schedule_log_to_main

        self.tab_widget = QTabWidget()
        self.tab_widget.addTab(self.table_filter_tab, "📊 数据发送")
        self.tab_widget.addTab(self.schedule_tab, "⏰ 定时发送")
        self.setCentralWidget(self.tab_widget)

    def _schedule_log_to_main(self, message):
        try:
            self.table_filter_tab.log(message)
        except Exception:
            pass

    def _schedule_dispatch_log(self, message):
        # 关键修复：该回调在 ScheduleDispatcher 调度线程里被直接调用，
        # 绝不能在这里碰 QTextEdit 等 QWidget（跨线程操作会与主线程
        # 重绘/写入竞争，导致界面卡死"无响应"甚至崩溃）。
        # 统一走 _schedule_log_signal 投递到主线程再写面板。
        self._post_schedule_log_from_worker(message)

    def _on_schedule_triggered(self, task: ScheduleTask, slot: str):
        # 调度器在子线程触发，通过信号投递到主线程创建 QThread。
        # 关键：QThread 对象必须在主线程创建（affinity = 主线程），
        # 否则 finished 信号和 deleteLater 无法被主线程事件循环处理，
        # 导致 QThread C++ 资源泄漏甚至主线程事件循环卡死。
        self._schedule_trigger_signal.emit(task, slot)

    def _on_schedule_trigger_in_main(self, task: ScheduleTask, slot: str):
        """主线程 slot：创建 ScheduleSendWorker（affinity = 主线程），再交给 ScheduleRun 线程执行。"""
        worker = ScheduleSendWorker(
            recipients=list(task.recipients),
            message=task.message,
            chat_delay=task.chat_delay,
            send_interval=task.send_interval,
            log_callback=self._post_schedule_log_from_worker,
            default_city=task.default_city,
            minimize_after=getattr(task, "minimize_after", True),
            attachment=getattr(task, "attachment", "") or "",
        )
        self._schedule_worker = worker
        threading.Thread(
            target=self._run_schedule_task_serialized,
            args=(task, slot, worker),
            daemon=True,
            name=f"ScheduleRun-{task.id[:8]}",
        ).start()

    def _run_schedule_task_serialized(self, task: ScheduleTask, slot: str, worker: "ScheduleSendWorker"):
        with self._schedule_running_lock:
            # 按任务配置启动防锁定守护
            if getattr(task, "keep_unlocked", False):
                if not self.workstation_guard.is_running():
                    self.workstation_guard.start()
                    self._post_schedule_log_from_worker(
                        f"[定时] 任务「{task.name}」已开启防锁定守护"
                    )
            self._post_schedule_log_from_worker(
                f"[定时] 开始执行任务「{task.name}」 时间点 {slot} 接收人数 {len(task.recipients)}"
            )
            # worker 已在主线程创建（affinity = 主线程），这里只负责启动和等待

            finished = threading.Event()
            result = {"success": 0, "failed": 0, "failed_list": []}

            def on_finished(success, failed, failed_list):
                # DirectConnection 下本函数在 worker QThread 内同步执行；
                # 这里只做 dict 赋值 + Event.set()，不碰任何 GUI，线程安全。
                if finished.is_set():
                    # 兜底去重：超时路径已 set() 过则不再覆盖统计。
                    return
                result["success"] = int(success or 0)
                result["failed"] = int(failed or 0)
                result["failed_list"] = list(failed_list or [])
                finished.set()

            # finished_with_result 用 DirectConnection：
            #   on_finished 只做 dict 赋值 + Event.set()，不碰 GUI，
            #   在 worker 线程内同步执行，不依赖任何线程的事件循环。
            worker.finished_with_result.connect(
                on_finished, Qt.ConnectionType.DirectConnection)
            # finished 用默认连接（AutoConnection）：
            #   worker affinity = 主线程 → emit 在 worker 线程，receiver 在主线程
            #   → Qt 自动用 QueuedConnection → cleanup + deleteLater 在主线程事件循环执行。
            #   不再使用 DirectConnection（在正在退出的 worker 线程里 deleteLater
            #   是 Qt 明确警告的危险模式，会导致 QThread 内部状态混乱、主线程卡死）。
            worker.finished.connect(
                lambda: self._cleanup_main_schedule_worker_after(worker)
            )
            worker.start()

            total_timeout = max(60.0, 60.0 * (len(task.recipients) or 1) * 10.0)
            total_timeout = min(total_timeout, 12 * 3600.0)
            finished.wait(timeout=total_timeout)
            if not finished.is_set():
                self._post_schedule_log_from_worker(
                    f"[定时] ⚠ 任务「{task.name}」执行超时（{int(total_timeout)}秒），强制停止"
                )
                try:
                    worker.stop()
                except Exception:
                    pass
            self._post_schedule_log_from_worker(
                f"[定时] 任务「{task.name}」时间点 {slot} 完成："
                f"成功{result['success']}，失败{result['failed']}"
            )
            # 发送完成后：按任务配置决定是否锁定电脑
            if getattr(task, "relock_after", False):
                try:
                    self.workstation_guard.maybe_relock_after_task(
                        self.schedule_store
                    )
                except Exception:
                    pass

    def _post_schedule_log_from_worker(self, message):
        # 关键修复：用 pyqtSignal 跨线程投递，不要用 QTimer.singleShot(0,...)
        # 因为 ScheduleSendWorker.run() 所在的 QThread 没有 Qt event loop，
        # singleShot 注册的 timer 永远不会触发；signal.emit 会通过 QueuedConnection
        # 自动投递到 receiver (MainWindow) 所在的主线程 event loop。
        try:
            self._schedule_log_signal.emit(str(message))
        except Exception:
            pass

    def _on_schedule_log_signal(self, message):
        # 主线程 slot：只写 schedule_tab.log 一个入口。
        # ScheduleTab.log 内部会通过 log_callback 转发到数据发送主面板，
        # 这里再直写 table_filter_tab 会导致主面板每条日志出现两行。
        try:
            self.schedule_tab.log(message)
        except Exception:
            pass

    def _cleanup_main_schedule_worker_after(self, worker):
        try:
            if worker is self._schedule_worker:
                self._schedule_worker = None
        finally:
            try:
                worker.deleteLater()
            except Exception:
                pass

    def _schedule_log_to_ui(self, message):
        try:
            self.schedule_tab.log(message)
        except Exception:
            pass

    def init_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        tray_icon = QApplication.instance().windowIcon()
        if tray_icon.isNull():
            tray_icon = self.style().standardIcon(
                QStyle.StandardPixmap.SP_ComputerIcon
            )
        self.setWindowIcon(tray_icon)
        self.tray_icon.setIcon(tray_icon)
        self.tray_icon.setToolTip(APP_TITLE)

        self.tray_menu = QMenu(self)
        self.show_window_action = QAction("显示主界面", self)
        self.show_window_action.triggered.connect(self.show_main_window)
        self.tray_menu.addAction(self.show_window_action)
        self.tray_menu.addSeparator()

        self.autostart_action = QAction("开机自启动", self)
        self.autostart_action.setCheckable(True)
        self.autostart_action.setChecked(is_autostart_enabled())
        self.autostart_action.toggled.connect(self._on_autostart_toggled)
        self.tray_menu.addAction(self.autostart_action)
        self.tray_menu.addSeparator()

        self.exit_action = QAction("退出程序", self)
        self.exit_action.triggered.connect(self.request_exit)
        self.tray_menu.addAction(self.exit_action)

        self.tray_icon.setContextMenu(self.tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        self._tray_available = QSystemTrayIcon.isSystemTrayAvailable()
        if self._tray_available:
            self.tray_icon.show()

    def log(self, message):
        self.table_filter_tab.log(message)

    def _on_autostart_toggled(self, checked: bool):
        ok = set_autostart(checked)
        if ok:
            self.log(f"开机自启动已{'开启' if checked else '关闭'}")
        else:
            # 写注册表失败，恢复 UI 勾选状态
            self.autostart_action.blockSignals(True)
            self.autostart_action.setChecked(not checked)
            self.autostart_action.blockSignals(False)
            QMessageBox.warning(self, "设置失败", "无法修改开机自启动设置，请检查权限。")

    def show_main_window(self):
        if self._closing_requested:
            return
        self.showNormal()
        self.raise_()
        self.activateWindow()

    def on_tray_activated(self, reason):
        if reason in (
            QSystemTrayIcon.ActivationReason.Trigger,
            QSystemTrayIcon.ActivationReason.DoubleClick,
        ):
            self.show_main_window()

    def closeEvent(self, event):
        if self._close_ready:
            self.tray_icon.hide()
            event.accept()
            return

        if self._closing_requested:
            event.ignore()
            return

        if self._tray_available:
            event.ignore()
            self.hide()
            if not self._tray_hint_shown:
                self.tray_icon.showMessage(
                    "程序仍在运行",
                    "程序已隐藏到系统托盘。右键托盘图标可退出程序。",
                    QSystemTrayIcon.MessageIcon.Information,
                    3000,
                )
                self._tray_hint_shown = True
            return

        event.ignore()
        self.request_exit()

    def request_exit(self):
        if self._closing_requested:
            return

        reply = QMessageBox.question(
            None,
            "确认退出",
            "确定要退出程序吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self._closing_requested = True
            self._shutdown_poll_count = 0
            self.setEnabled(False)
            self.setWindowTitle("正在安全退出，请稍候...")
            self._request_worker_stop()
            QTimer.singleShot(100, self._poll_worker_shutdown)

    def _request_worker_stop(self):
        worker = self.table_filter_tab.worker
        if worker and worker.isRunning():
            worker.set_stopped(True)
            worker.set_paused(False)

        excel_worker = self.table_filter_tab.excel_worker
        if excel_worker and excel_worker.isRunning():
            excel_worker.stop()

        schedule_worker = getattr(self, "_schedule_worker", None)
        if schedule_worker and isinstance(schedule_worker, QThread):
            try:
                if schedule_worker.isRunning():
                    schedule_worker.stop()
            except Exception:
                pass

        dispatcher = getattr(self, "schedule_dispatcher", None)
        if dispatcher is not None:
            try:
                dispatcher.stop()
            except Exception:
                pass

        # 停止电脑锁定守护并恢复屏保/锁屏策略原状
        guard = getattr(self, "workstation_guard", None)
        if guard is not None:
            try:
                guard.stop()
            except Exception:
                pass

    def _has_running_workers(self):
        worker = self.table_filter_tab.worker
        excel_worker = self.table_filter_tab.excel_worker
        schedule_worker = getattr(self, "_schedule_worker", None)
        running = bool(
            (worker and worker.isRunning())
            or (excel_worker and excel_worker.isRunning())
        )
        if not running and schedule_worker is not None:
            try:
                if isinstance(schedule_worker, QThread) and schedule_worker.isRunning():
                    running = True
            except Exception:
                pass
        return running

    def _poll_worker_shutdown(self):
        if not self._closing_requested:
            return

        if not self._has_running_workers():
            self._close_ready = True
            self.close()
            QApplication.instance().quit()
            return

        self._shutdown_poll_count += 1
        if self._shutdown_poll_count >= 150:
            self.setEnabled(True)
            reply = QMessageBox.question(
                self,
                "后台操作仍在结束",
                "后台操作尚未结束，是否继续等待安全退出？\n"
                "选择“否”将返回程序，不会强制终止线程。",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes
            )
            if reply == QMessageBox.StandardButton.No:
                self._closing_requested = False
                self.setWindowTitle(APP_TITLE)
                self.show_main_window()
                return

            self._shutdown_poll_count = 0
            self.setEnabled(False)

        QTimer.singleShot(100, self._poll_worker_shutdown)


def run_gui():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setQuitOnLastWindowClosed(False)
    app.setStyle("Fusion")
    
    app_icon = get_application_icon()
    if not app_icon.isNull():
        app.setWindowIcon(app_icon)

    instance_lock = acquire_instance_lock()
    if instance_lock is None:
        QMessageBox.information(
            None,
            "程序正在运行",
            f"{APP_TITLE} 已经在运行，请从系统托盘打开。",
        )
        return
    app.instance_lock = instance_lock
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())
