import sys
import os
import time
import threading
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QLineEdit, QPushButton, QTextEdit,
    QComboBox, QListWidget, QListWidgetItem, QGroupBox,
    QCheckBox, QProgressBar, QMessageBox, QSplitter, QSpinBox,
    QDoubleSpinBox, QDialog, QDialogButtonBox, QFileDialog
)
from PyQt6.QtCore import Qt, pyqtSignal, QObject, QThread
from PyQt6.QtGui import QFont, QIcon

from modules.wechat_sender import WeChatSender
from modules.table_processor import TableProcessor
from modules.wps_extractor import WPSExtractor



class WorkerSignals(QObject):
    finished = pyqtSignal()
    error = pyqtSignal(str)
    progress = pyqtSignal(int)
    result = pyqtSignal(object)
    log = pyqtSignal(str)





class ExtractionWorker(QThread):
    def __init__(self, document_url, extraction_type, **kwargs):
        super().__init__()
        self.document_url = document_url
        self.extraction_type = extraction_type
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    def run(self):
        try:
            self.signals.log.emit("开始提取文档内容...")
            self.signals.log.emit(f"URL: {self.document_url[:50]}...")
            self.signals.log.emit(f"类型: {self.extraction_type}")
            self.signals.log.emit(f"参数: {str(self.kwargs)[:100]}...")
            
            result = WPSExtractor.extract_from_document(
                self.document_url,
                extraction_type=self.extraction_type,
                **self.kwargs
            )
            
            if result is None:
                self.signals.log.emit("⚠ 提取返回空结果")
            elif isinstance(result, list):
                self.signals.log.emit(f"✓ 提取成功: {len(result)} 个表格")
            else:
                self.signals.log.emit(f"✓ 提取成功: {len(str(result))} 字符")
                
            self.signals.result.emit(result)
            self.signals.log.emit("文档内容提取完成")
        except Exception as e:
            import traceback
            error_msg = f"提取失败: {str(e)}"
            self.signals.log.emit(f"❌ {error_msg}")
            self.signals.log.emit(f"详细错误: {traceback.format_exc()[:500]}")
            self.signals.error.emit(error_msg)
        finally:
            self.signals.finished.emit()


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
        finally:
            self.signals.finished.emit()

    def _read_excel(self, file_path):
        import pandas as pd
        
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
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
                ws = wb[sheet_name]
                
                data = []
                for row in ws.iter_rows(values_only=True):
                    data.append([str(cell) if cell is not None else "" for cell in row])
                
                wb.close()
                
                if len(data) > 0:
                    headers = [h if str(h).strip() != "" else f"列{i+1}" for i, h in enumerate(data[0])]
                    self.signals.log(f"表头(openpyxl): {headers}")
                    return headers, data
                else:
                    raise Exception("Sheet为空")
                    
            except Exception as e2:
                raise Exception(f"openpyxl读取也失败: {e2}")


class SendWorker(QThread):
    def __init__(self, tasks, send_interval=2, chat_delay=0.8):
        super().__init__()
        self.tasks = tasks
        self.send_interval = send_interval
        self.chat_delay = chat_delay
        self.signals = WorkerSignals()
        self.paused_event = threading.Event()
        self.stopped_event = threading.Event()

    def run(self):
        success_count = 0
        failed_count = 0
        failed_tasks = []
        total_count = len(self.tasks)
        
        try:
            self.signals.log.emit("初始化微信客户端...")
            sender = WeChatSender()
            if not sender.initialize():
                self.signals.error.emit("微信未登录或未打开")
                return
            self.signals.log.emit("微信客户端初始化成功")
            
            for i, task in enumerate(self.tasks):
                if self.stopped_event.is_set():
                    self.signals.log.emit("⏹ 发送已停止")
                    break
                
                while self.paused_event.is_set() and not self.stopped_event.is_set():
                    time.sleep(0.1)
                if self.stopped_event.is_set():
                    self.signals.log.emit("⏹ 发送已停止")
                    break
                
                name, person_data, recipient, custom_msg = task
                
                try:
                    if i == 0:
                        self.signals.log.emit("等待微信就绪...")
                        time.sleep(0.2)
                    
                    self.signals.log.emit(f"[{i+1}/{total_count}] 正在发送给 {recipient} ({name})...")
                    self.signals.progress.emit(int((i + 1) / total_count * 100))
                    
                    messages = []
                    max_message_length = 2000
                    for j in range(0, len(person_data), max_message_length):
                        messages.append(person_data[j:j+max_message_length])
                    
                    success = sender.send_multiple_messages(messages, recipient, chat_delay=self.chat_delay)
                    if success:
                        self.signals.log.emit(f"[{i+1}/{total_count}] ✅ 成功发送给 {recipient}")
                        success_count += 1
                    else:
                        self.signals.log.emit(f"[{i+1}/{total_count}] ❌ 发送失败: {recipient}")
                        failed_count += 1
                        failed_tasks.append((name, person_data, recipient, custom_msg))
                    
                    if custom_msg:
                        sender.send_message(custom_msg, recipient, chat_delay=self.chat_delay, fast_mode=True)
                        self.signals.log.emit(f"[{i+1}/{total_count}] 已发送自定义消息")
                    
                except Exception as e:
                    self.signals.log.emit(f"[{i+1}/{total_count}] ❌ 发送异常: {name} - {str(e)}")
                    failed_count += 1
                    failed_tasks.append((name, person_data, recipient, custom_msg))
                
                if i < total_count - 1 and not self.stopped_event.is_set():
                    time.sleep(self.send_interval)
            
            if failed_tasks:
                self.signals.log.emit(f"\n--- 发送失败列表 ({failed_count}人) ---")
                for name, _, recipient, _ in failed_tasks:
                    self.signals.log.emit(f"❌ {name} → {recipient}")
            
            self.signals.log.emit(f"\n发送完成！成功: {success_count}, 失败: {failed_count}, 总计: {total_count}")
            self.signals.result.emit((success_count, failed_count, total_count, failed_tasks))
            
        except Exception as e:
            import traceback
            self.signals.log.emit(f"❌ 发送线程异常: {str(e)}")
            self.signals.log.emit(f"详细错误: {traceback.format_exc()[:300]}")
            self.signals.error.emit(str(e))

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


class APIAuthWorker(QThread):
    """后台线程：处理WPS API授权"""
    def __init__(self, app_id, app_key, redirect_uri):
        super().__init__()
        self.app_id = app_id
        self.app_key = app_key
        self.redirect_uri = redirect_uri
        self.signals = WorkerSignals()

    def run(self):
        try:
            from modules.wps_official_api import WPSOfficialAPI

            self.signals.log.emit("🔐 开始WPS开放平台授权流程...")
            self.signals.log.emit(f"AppID: {self.app_id}")
            self.signals.log.emit(f"回调地址: {self.redirect_uri}")
            self.signals.log.emit("即将打开浏览器进行授权，请在浏览器中完成登录和授权...")
            self.signals.log.emit("如果浏览器没有自动打开，请查看终端输出的授权URL")

            api = WPSOfficialAPI(self.app_id, self.app_key, self.redirect_uri)
            token_data = api.authorize()

            if token_data and token_data.get("access_token"):
                self.signals.log.emit("✓ 授权成功！")
                self.signals.log.emit(f"  access_token: {token_data['access_token'][:20]}...")
                self.signals.log.emit(f"  有效期: {token_data.get('expires_in', '未知')} 秒")
                self.signals.result.emit(token_data)
            else:
                self.signals.log.emit("✗ 授权失败：未获取到 access_token")
                self.signals.error.emit("未获取到 access_token")

        except Exception as e:
            import traceback
            error_msg = f"授权出错: {str(e)}"
            self.signals.log.emit(f"✗ {error_msg}")
            self.signals.log.emit(f"详细错误: {traceback.format_exc()[:500]}")
            self.signals.error.emit(error_msg)
        finally:
            self.signals.finished.emit()


class APIExtractWorker(QThread):
    """后台线程：使用WPS官方API提取表格数据"""
    def __init__(self, app_id, app_key, redirect_uri, document_url, sheet_name):
        super().__init__()
        self.app_id = app_id
        self.app_key = app_key
        self.redirect_uri = redirect_uri
        self.document_url = document_url
        self.sheet_name = sheet_name
        self.signals = WorkerSignals()

    def run(self):
        try:
            from modules.wps_official_api import WPSOfficialAPI

            self.signals.log.emit("📊 使用WPS官方API提取表格数据...")
            self.signals.log.emit(f"URL: {self.document_url}")
            self.signals.log.emit(f"Sheet: {self.sheet_name or '(未指定)'}")

            api = WPSOfficialAPI(self.app_id, self.app_key, self.redirect_uri)
            result = api.extract_table(self.document_url, self.sheet_name)

            if result and len(result) > 0:
                table_data = result[0]
                self.signals.log.emit(f"✓ API提取成功: {len(table_data)} 行, {len(table_data[0])} 列")
                self.signals.log.emit(f"表头: {table_data[0]}")
                self.signals.result.emit(result)
            else:
                self.signals.log.emit("✗ API提取失败")
                self.signals.result.emit(None)

        except Exception as e:
            import traceback
            error_msg = f"API提取出错: {str(e)}"
            self.signals.log.emit(f"✗ {error_msg}")
            self.signals.log.emit(f"详细错误: {traceback.format_exc()[:500]}")
            self.signals.error.emit(error_msg)
        finally:
            self.signals.finished.emit()


class TableFilterTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.table_data = None
        self.processor = None
        self.wechat_mapping = {}
        self.filter_conditions = []
        
        self.init_ui()
        self.connect_signals()

    def init_ui(self):
        main_layout = QVBoxLayout()
        
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
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
        
        interval_layout = QHBoxLayout()
        interval_layout.addWidget(QLabel("发送间隔(秒):"))
        self.send_interval_spin = QSpinBox()
        self.send_interval_spin.setRange(0, 30)
        self.send_interval_spin.setValue(2)
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
        
        right_layout.addWidget(control_group)
        
        right_layout.addStretch()
        splitter.addWidget(right_panel)
        
        splitter.setSizes([285, 280, 285])
        
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

    def connect_signals(self):
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

    def open_extract_columns_dialog(self):
        if not hasattr(self, 'headers') or not self.headers:
            return
        
        dialog = MultiSelectDialog("选择要提取的列", self.headers, self.selected_columns, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.selected_columns = dialog.selected_columns
            if self.selected_columns:
                if len(self.selected_columns) <= 3:
                    self.selected_columns_label.setText(f"已选择: {', '.join(self.selected_columns)}")
                else:
                    self.selected_columns_label.setText(f"已选择: {len(self.selected_columns)}列")
            else:
                self.selected_columns_label.setText("已选择: 0列")

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
        
        self.log(f"筛选后找到 {len(persons)} 个人")

    def on_person_selection(self):
        selected_items = self.persons_list.selectedItems()
        if selected_items and self.table_data:
            self.send_btn.setEnabled(True)
            self.start_send_btn.setEnabled(True)
            self.preview_selected_data()
        else:
            self.send_btn.setEnabled(False)
            self.start_send_btn.setEnabled(False)

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
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls)"
        )
        
        if file_path:
            self.current_excel_path = file_path
            self.current_file_label.setText(f"当前文件: {os.path.basename(file_path)}")
            self.open_excel_btn.setEnabled(True)
            self.reload_btn.setEnabled(True)
            
            self._load_excel_data(file_path)
    
    def reload_excel_file(self):
        if hasattr(self, 'current_excel_path') and self.current_excel_path:
            self._load_excel_data(self.current_excel_path)
    
    def open_current_excel(self):
        if hasattr(self, 'current_excel_path') and self.current_excel_path:
            try:
                os.startfile(self.current_excel_path)
                self.log(f"已打开文件: {self.current_excel_path}")
            except Exception as e:
                self.log(f"✗ 打开文件失败: {e}")
                QMessageBox.warning(self, "打开失败", f"无法打开文件: {e}")
    
    def _load_excel_data(self, file_path):
        self.log(f"选择Excel文件: {file_path}")
        
        self.sheet_combo.setEnabled(False)
        self.name_column_combo.setEnabled(False)
        self.extract_columns_btn.setEnabled(False)
        self.wechat_column_combo.setEnabled(False)
        self.load_data_btn.setEnabled(False)
        
        self.excel_worker = ExcelReadWorker(file_path)
        self.excel_worker.signals.result.connect(self.on_excel_read_result)
        self.excel_worker.signals.error.connect(self.on_excel_read_error)
        self.excel_worker.signals.finished.connect(self.on_excel_read_finished)
        self.excel_worker.signals.log.connect(self.log)
        self.excel_worker.start()
    
    def on_excel_read_result(self, result):
        self.sheet_names = result['sheet_names']
        self.current_sheet = result['current_sheet']
        self.headers = result['headers']
        self.table_data = result['data']
        
        self.sheet_combo.clear()
        self.sheet_combo.addItems(self.sheet_names)
        
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.setCurrentText(self.current_sheet)
        self.sheet_combo.blockSignals(False)
        
        self.sheet_combo.setEnabled(True)
        self._update_column_combos()
        self.log("Excel文件读取完成！请选择Sheet和列，然后点击'加载数据'")
    
    def on_excel_read_error(self, error):
        QMessageBox.warning(self, "读取失败", error)
        self.excel_btn.setEnabled(True)
    
    def on_excel_read_finished(self):
        if hasattr(self, 'excel_worker'):
            if self.excel_worker.isRunning():
                self.excel_worker.stop()
                self.excel_worker.wait()
            self.excel_worker.deleteLater()
            del self.excel_worker
    
    def on_sheet_changed(self, index):
        if index >= 0 and hasattr(self, 'current_excel_path') and self.current_excel_path:
            sheet_name = self.sheet_combo.itemText(index)
            self.log(f"--- 切换到Sheet: {sheet_name} ---")
            
            self.sheet_combo.setEnabled(False)
            self.name_column_combo.setEnabled(False)
            self.extract_columns_btn.setEnabled(False)
            self.wechat_column_combo.setEnabled(False)
            self.load_data_btn.setEnabled(False)
            
            if hasattr(self, 'excel_worker') and self.excel_worker and self.excel_worker.isRunning():
                self.excel_worker.stop()
                self.excel_worker.wait()
            
            self.excel_worker = ExcelReadWorker(self.current_excel_path, sheet_name)
            self.excel_worker.signals.result.connect(self.on_excel_read_result)
            self.excel_worker.signals.error.connect(self.on_excel_read_error)
            self.excel_worker.signals.finished.connect(self.on_excel_read_finished)
            self.excel_worker.signals.log.connect(self.log)
            self.excel_worker.start()
    
    def _update_column_combos(self):
        self.name_column_combo.clear()
        self.name_column_combo.addItems(self.headers)
        
        self.wechat_column_combo.clear()
        self.wechat_column_combo.addItems([""] + self.headers)
        
        self.selected_columns = []
        self.selected_columns_label.setText("已选择: 0列")
        
        self.name_column_combo.setEnabled(True)
        self.extract_columns_btn.setEnabled(True)
        self.wechat_column_combo.setEnabled(True)
        self.load_data_btn.setEnabled(True)
    
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
        
        max_lines = 500
        if self.log_text.document().blockCount() > max_lines:
            cursor = self.log_text.textCursor()
            cursor.movePosition(cursor.MoveOperation.Start)
            cursor.movePosition(cursor.MoveOperation.NextBlock)
            cursor.select(cursor.SelectionType.BlockUnderCursor)
            cursor.removeSelectedText()
            cursor.deleteChar()

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
        
        tasks = []
        missing_wechat = []
        
        for item in selected_items:
            name = item.text()
            person_data = self.processor.get_person_data(
                name, 
                name_column, 
                extract_columns,
                self.filter_conditions
            )
            
            if not person_data:
                self.log(f"未找到 {name} 的数据，跳过")
                continue
            
            recipient = self.wechat_mapping.get(name, "")
            if not recipient:
                recipient = self.wechat_edit.text().strip()
            
            if not recipient:
                recipient = name
            
            tasks.append((name, person_data, recipient, custom_msg))
        
        if missing_wechat:
            QMessageBox.warning(self, "警告", f"以下人员缺少微信接收人，已跳过:\n{', '.join(missing_wechat)}")
        
        if not tasks:
            QMessageBox.warning(self, "警告", "没有可发送的任务")
            return
        
        max_display = 20
        confirm_text = f"即将向以下 {len(tasks)} 人发送消息:\n\n"
        for i, (name, _, recipient, _) in enumerate(tasks):
            if i >= max_display:
                confirm_text += f"  • ... 还有 {len(tasks) - max_display} 人\n"
                break
            confirm_text += f"  • {name} → {recipient}\n"
        confirm_text += f"\n发送间隔: {self.send_interval_spin.value()}秒"
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
        
        self.worker = SendWorker(
            tasks=tasks,
            send_interval=self.send_interval_spin.value(),
            chat_delay=self.chat_delay_spin.value()
        )
        self.worker.signals.result.connect(self.on_send_result)
        self.worker.signals.error.connect(self.on_send_error)
        self.worker.signals.finished.connect(self.on_send_finished)
        self.worker.signals.log.connect(self.log)
        self.worker.signals.progress.connect(self.on_send_progress)
        self.worker.start()

    def pause_send(self):
        if hasattr(self, 'worker') and self.worker.isRunning():
            if self.worker.is_paused():
                self.worker.set_paused(False)
                self.pause_send_btn.setText("⏸ 暂停")
                self.log("▶ 继续发送")
            else:
                self.worker.set_paused(True)
                self.pause_send_btn.setText("▶ 继续")
                self.log("⏸ 发送已暂停")
    
    def stop_send(self):
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.set_stopped(True)
            self.worker.set_paused(False)
            self.log("⏹ 正在停止发送...")

    def retry_send(self):
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
        
        self.worker = SendWorker(
            tasks=failed_tasks,
            send_interval=self.send_interval_spin.value(),
            chat_delay=self.chat_delay_spin.value()
        )
        self.worker.signals.result.connect(self.on_send_result)
        self.worker.signals.error.connect(self.on_send_error)
        self.worker.signals.finished.connect(self.on_send_finished)
        self.worker.signals.log.connect(self.log)
        self.worker.signals.progress.connect(self.on_send_progress)
        self.worker.start()

    def on_send_finished(self):
        self.send_btn.setEnabled(True)
        self.start_send_btn.setEnabled(True)
        self.pause_send_btn.setEnabled(False)
        self.stop_send_btn.setEnabled(False)
        self.pause_send_btn.setText("⏸ 暂停")
        
        if hasattr(self, 'worker'):
            self.worker.wait()
            self.worker.deleteLater()
            del self.worker

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


class SimpleModeTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.extracted_content = None
        
        self.init_ui()
        self.connect_signals()

    def init_ui(self):
        main_layout = QVBoxLayout()
        
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        url_group = QGroupBox("文档设置")
        url_layout = QVBoxLayout(url_group)
        
        url_layout.addWidget(QLabel("WPS在线文档URL:"))
        self.url_edit = QLineEdit()
        self.url_edit.setPlaceholderText("https://www.wps.cn/xxx")
        url_layout.addWidget(self.url_edit)
        
        url_layout.addWidget(QLabel("提取类型:"))
        self.extract_type_combo = QComboBox()
        self.extract_type_combo.addItems(["全部文本", "关键词", "表格", "CSS选择器"])
        url_layout.addWidget(self.extract_type_combo)
        
        self.keyword_group = QGroupBox("关键词设置")
        keyword_layout = QVBoxLayout(self.keyword_group)
        
        keyword_layout.addWidget(QLabel("关键词:"))
        self.keyword_edit = QLineEdit()
        keyword_layout.addWidget(self.keyword_edit)
        
        layout_row = QHBoxLayout()
        layout_row.addWidget(QLabel("前后字符数:"))
        self.before_chars = QSpinBox()
        self.before_chars.setMaximum(500)
        self.before_chars.setValue(0)
        layout_row.addWidget(self.before_chars)
        layout_row.addWidget(QLabel("后:"))
        self.after_chars = QSpinBox()
        self.after_chars.setMaximum(2000)
        self.after_chars.setValue(100)
        layout_row.addWidget(self.after_chars)
        keyword_layout.addLayout(layout_row)
        
        url_layout.addWidget(self.keyword_group)
        
        self.selector_group = QGroupBox("CSS选择器")
        selector_layout = QVBoxLayout(self.selector_group)
        selector_layout.addWidget(QLabel("选择器:"))
        self.selector_edit = QLineEdit()
        selector_layout.addWidget(self.selector_edit)
        url_layout.addWidget(self.selector_group)
        
        self.extract_btn = QPushButton("提取内容")
        self.extract_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 8px;")
        url_layout.addWidget(self.extract_btn)
        
        left_layout.addWidget(url_group)
        
        send_group = QGroupBox("发送设置")
        send_layout = QVBoxLayout(send_group)
        
        send_layout.addWidget(QLabel("微信接收人:"))
        self.recipient_edit = QLineEdit()
        send_layout.addWidget(self.recipient_edit)
        
        self.send_btn = QPushButton("发送到微信")
        self.send_btn.setStyleSheet("background-color: #2196F3; color: white; padding: 8px;")
        self.send_btn.setEnabled(False)
        send_layout.addWidget(self.send_btn)
        
        left_layout.addWidget(send_group)
        
        left_layout.addStretch()
        splitter.addWidget(left_panel)
        
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        preview_group = QGroupBox("提取内容预览")
        preview_layout = QVBoxLayout(preview_group)
        
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        preview_layout.addWidget(self.preview_text)
        
        right_layout.addWidget(preview_group)
        
        log_group = QGroupBox("日志")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_text)
        
        right_layout.addWidget(log_group)
        
        right_layout.addStretch()
        splitter.addWidget(right_panel)
        
        splitter.setSizes([350, 650])
        
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)
        
        self.update_group_visibility()

    def connect_signals(self):
        self.extract_btn.clicked.connect(self.extract_content)
        self.send_btn.clicked.connect(self.send_content)
        self.extract_type_combo.currentIndexChanged.connect(self.update_group_visibility)

    def update_group_visibility(self):
        extract_type = self.extract_type_combo.currentText()
        self.keyword_group.setVisible(extract_type == "关键词")
        self.selector_group.setVisible(extract_type == "CSS选择器")

    def log(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())
        
        max_lines = 500
        if self.log_text.document().blockCount() > max_lines:
            cursor = self.log_text.textCursor()
            cursor.movePosition(cursor.MoveOperation.Start)
            cursor.movePosition(cursor.MoveOperation.NextBlock)
            cursor.select(cursor.SelectionType.BlockUnderCursor)
            cursor.removeSelectedText()
            cursor.deleteChar()

    def extract_content(self):
        url = self.url_edit.text().strip()
        if not url:
            QMessageBox.warning(self, "警告", "请输入文档URL")
            return
        
        extract_type = self.extract_type_combo.currentText()
        
        type_map = {
            "全部文本": "all",
            "关键词": "keyword",
            "表格": "table",
            "CSS选择器": "selector"
        }
        
        extraction_type = type_map.get(extract_type, "all")
        kwargs = {}
        
        if extraction_type == "keyword":
            keyword = self.keyword_edit.text().strip()
            if not keyword:
                QMessageBox.warning(self, "警告", "请输入关键词")
                return
            kwargs["keyword"] = keyword
            kwargs["before_chars"] = self.before_chars.value()
            kwargs["after_chars"] = self.after_chars.value()
        
        if extraction_type == "selector":
            selector = self.selector_edit.text().strip()
            if not selector:
                QMessageBox.warning(self, "警告", "请输入CSS选择器")
                return
            kwargs["selector"] = selector
        
        self.log(f"正在提取{extract_type}...")
        self.extract_btn.setEnabled(False)
        
        self.worker = ExtractionWorker(
            document_url=url,
            extraction_type=extraction_type,
            **kwargs
        )
        self.worker.signals.result.connect(self.on_extract_result)
        self.worker.signals.error.connect(self.on_extract_error)
        self.worker.signals.finished.connect(lambda: self.extract_btn.setEnabled(True))
        self.worker.signals.log.connect(self.log)
        self.worker.start()

    def on_extract_result(self, result):
        self.extracted_content = result
        
        if not result:
            QMessageBox.warning(self, "警告", "未提取到内容")
            return
        
        if isinstance(result, list):
            if isinstance(result[0], list):
                formatted = "\n".join([" | ".join(row) for table in result for row in table])
            else:
                formatted = "\n\n".join(result)
        else:
            formatted = str(result)
        
        self.preview_text.setPlainText(formatted)
        self.send_btn.setEnabled(True)
        self.log(f"提取完成，内容长度: {len(formatted)}")

    def on_extract_error(self, error):
        QMessageBox.critical(self, "错误", f"提取失败: {error}")
        self.log(f"提取失败: {error}")

    def send_content(self):
        if not self.extracted_content:
            QMessageBox.warning(self, "警告", "请先提取内容")
            return
        
        recipient = self.recipient_edit.text().strip()
        if not recipient:
            QMessageBox.warning(self, "警告", "请输入微信接收人")
            return
        
        content = self.preview_text.toPlainText()
        
        self.log(f"正在发送消息给 {recipient}...")
        self.send_btn.setEnabled(False)
        
        self.worker = SendWorker([("", content, recipient)], send_interval=0)
        self.worker.signals.error.connect(self.on_send_error)
        self.worker.signals.finished.connect(lambda: self.send_btn.setEnabled(True))
        self.worker.signals.log.connect(self.log)
        self.worker.start()

    def on_send_error(self, error):
        QMessageBox.critical(self, "错误", f"发送失败: {error}")
        self.log(f"发送失败: {error}")




class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("表格自动发送By春风予Lu")
        self.setGeometry(50, 50, 850, 580)
        
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout(central_widget)
        
        tab_widget = QTabWidget()
        
        self.table_filter_tab = TableFilterTab()
        tab_widget.addTab(self.table_filter_tab, "表格筛选模式")
        
        self.simple_tab = SimpleModeTab()
        tab_widget.addTab(self.simple_tab, "简单提取模式")
        
        layout.addWidget(tab_widget)

    def log(self, message):
        current_tab_index = self.findChild(QTabWidget).currentIndex()
        if current_tab_index == 0:
            self.table_filter_tab.log(message)
        else:
            self.simple_tab.log(message)

    def closeEvent(self, event):
        reply = QMessageBox.question(
            self, "确认退出", "确定要退出程序吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            if hasattr(self, 'table_filter_tab'):
                if hasattr(self.table_filter_tab, 'worker') and self.table_filter_tab.worker.isRunning():
                    self.table_filter_tab.worker.set_stopped(True)
                    self.table_filter_tab.worker.set_paused(False)
                    self.table_filter_tab.worker.wait()
                
                if hasattr(self.table_filter_tab, 'excel_worker') and self.table_filter_tab.excel_worker.isRunning():
                    self.table_filter_tab.excel_worker.stop()
                    self.table_filter_tab.excel_worker.wait()
            
            event.accept()
        else:
            event.ignore()


def run_gui():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    import os
    icon_path = os.path.join(os.path.dirname(__file__), "..", "love.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())