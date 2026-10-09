import sys
import os
import glob
import logging
import threading
import time
from typing import List, Optional, Set, Dict, Tuple, Any
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit,
    QComboBox, QListWidget, QListWidgetItem, QGroupBox,
    QCheckBox, QProgressBar, QMessageBox, QSplitter, QTabWidget,
    QDoubleSpinBox, QDialog, QDialogButtonBox, QFileDialog,
    QMenu, QStyle, QSystemTrayIcon, QSpinBox, QTimeEdit, QDateEdit,
    QStackedWidget,
    QLayout, QRadioButton, QButtonGroup, QFrame,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QPlainTextEdit,
)
from PyQt6.QtCore import (
    Qt, pyqtSignal, QObject, QThread, QTimer, QLockFile, QStandardPaths,
    QTime, QDate, QEvent, QSize, QRect, QPoint, QPropertyAnimation, QEasingCurve,
    QAbstractAnimation, QVariantAnimation, QUrl,
)
from PyQt6.QtGui import QAction, QFont, QIcon, QGuiApplication, QDesktopServices

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
from modules.monitor_config import MonitorManager, MonitorTask
from modules.monitor_worker import MonitorWorker
from modules import wps_cloud
from modules.wps_cloud import (
    WpsOAuthStore,
    WpsCloudClient,
    WpsCloudError,
    WpsAuthExpired,
    parse_file_input,
    new_temp_xlsx_path,
    cleanup_temp_file,
    read_all_sheets,
)
from modules import updater
from modules.updater import (
    APP_VERSION,
    UpdateCheckWorker,
    UpdateDialog,
    UpdateState,
)


# gui 模块自己的日志（Toast、定时触发/完成、跨线程信号等关键路径均需可追溯）
logger = logging.getLogger(__name__)


APP_TITLE = "表格自动发送By春风予Lu"
INSTANCE_LOCK_NAME = "ExcelSendWx.lock"
# 与 installer.iss 的 Run 键值名一致，确保安装器勾选和托盘菜单勾选操作同一注册表项
_AUTOSTART_REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
_AUTOSTART_REG_NAME = "ExcelSendWx"


def _app_is_dark() -> bool:
    """根据当前应用调色板判断是否为深色（夜间）模式。"""
    try:
        window_lightness = QApplication.palette().color(
            QApplication.palette().ColorRole.Window
        ).lightness()
        return window_lightness < 128
    except Exception:
        return False


def _chip_css() -> str:
    """芯片按钮样式：胶囊形，选中态高亮；明/暗主题均清晰可读。"""
    if _app_is_dark():
        return (
            "QPushButton{text-align:center;background:#33363b;color:#d8dae0;"
            "border:1px solid #4a4d55;padding:4px 14px;font-size:12px;"
            "border-radius:13px}"
            "QPushButton:hover{background:#3d4148;border-color:#5d636e}"
            "QPushButton:checked{background:#3a6ea5;border-color:#5a9bd6;"
            "color:#ffffff;font-weight:bold}"
            "QPushButton:checked:hover{background:#467bb8}"
        )
    return (
        "QPushButton{text-align:center;background:#ffffff;color:#333333;"
        "border:1px solid #c9ccd4;padding:4px 14px;font-size:12px;"
        "border-radius:13px}"
        "QPushButton:hover{background:#eef1f6;border-color:#aeb4c0}"
        "QPushButton:checked{background:#2f78c4;border-color:#2f78c4;"
        "color:#ffffff;font-weight:bold}"
        "QPushButton:checked:hover{background:#3b85d1}"
    )


def _compact_btn_css() -> str:
    """全局精简/详细按钮样式：明暗主题适配。"""
    if _app_is_dark():
        return (
            "QPushButton{padding:3px 10px;font-size:12px;color:#f2f2f2;"
            "border:1px solid #4a4d55;border-radius:3px;background:#3a3d44}"
            "QPushButton:hover{background:#494d55}"
        )
    return (
        "QPushButton{padding:3px 10px;font-size:12px;color:#1f1f1f;"
        "border:1px solid #bbb;border-radius:3px;background:#f5f5f5}"
        "QPushButton:hover{background:#e8e8e8}"
    )


class ElasticPage(QWidget):
    """弹性 tab 页面容器：非当前显示的页不报告尺寸提示。

    QTabWidget 内部 QStackedLayout 枚举所有页面的 sizeHint/minimumSizeHint
    （不跳过隐藏页），导致当前页分组全部折叠后，窗口仍被另一个未折叠的 tab
    页顶住而无法收缩。本容器在非当前页时返回 (0,0)，当前页时透传布局真实
    尺寸；切换 tab 后窗口会按新页重新弹性适应。
    """

    def _is_noncurrent(self) -> bool:
        p = self.parent()
        # QTabWidget 内部用 QStackedWidget 承载页面
        while p is not None:
            if isinstance(p, QStackedWidget):
                return p.currentWidget() is not self
            p = p.parent()
        return False

    def sizeHint(self):
        if self._is_noncurrent():
            return QSize(0, 0)
        return super().sizeHint()

    def minimumSizeHint(self):
        if self._is_noncurrent():
            return QSize(0, 0)
        return super().minimumSizeHint()

    def content_widget(self) -> QWidget:
        """返回容器内承载的真实 tab 页面（ElasticPage 只是弹性包装层）。"""
        lay = self.layout()
        if lay is not None and lay.count():
            w = lay.itemAt(0).widget()
            if w is not None:
                return w
        return self


class _ElasticTabWidget(QTabWidget):
    """弹性 TabWidget（配套 ElasticPage 使用，保留 tab 基础行为）。"""
    pass


class ChipSection(QWidget):
    """芯片面板：无独立标题，由所属 ChipBar 上的芯片按钮控制果冻展开/折叠。

    与 ChipBar 配套使用::

        bar = ChipBar(exclusive=True)
        parent_layout.addWidget(bar)
        section = bar.add_section("发送设置")
        parent_layout.addWidget(section)
        section.contentLayout().addWidget(...)
    """

    collapsedChanged = pyqtSignal(bool)  # True = 已折叠

    def __init__(self, parent=None, collapsed: bool = False):
        super().__init__(parent)
        self._collapsed = False
        self._title_text = ""
        self._chip: Optional[QPushButton] = None
        self._chip_bar = None
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self._content = QWidget()
        self._content_layout = QVBoxLayout(self._content)
        self._content_layout.setContentsMargins(8, 6, 8, 8)
        self._content_layout.setSpacing(5)
        outer.addWidget(self._content)

        # 单一持久动画对象：反复折叠/展开时只 stop + 重设参数重启，
        # 不频繁创建/销毁动画对象（避免极端连点时 deleteLater 堆积引发原生崩溃）
        self._anim = QPropertyAnimation(self._content, b"maximumHeight", self)
        self._anim.finished.connect(self._on_anim_finished)
        self._anim_target_collapsed = False
        self._no_accordion_close = False
        if collapsed:
            self.set_collapsed(True)

    # ---- public API ----
    def contentLayout(self) -> QVBoxLayout:
        """返回内容区布局，供外部 addWidget / addLayout。"""
        return self._content_layout

    def set_title(self, title: str) -> None:
        self._title_text = title
        if self._chip is not None:
            self._chip.setText(title)

    def is_collapsed(self) -> bool:
        return self._collapsed

    def _kill_anim(self) -> None:
        """停止进行中的折叠动画并解除高度限制（stop 不触发 finished，
        可防止旧动画结束回调把内容显示状态写反）。"""
        try:
            self._anim.stop()
        except Exception:
            pass
        self._content.setMaximumHeight(16777215)

    def set_collapsed(self, collapsed: bool, animate: bool = False,
                      accordion_close: bool = True) -> None:
        """折叠/展开面板。

        :param accordion_close: 手风琴排中，程序自动展开时传 False 可避免
            挤掉用户正在查看/编辑的其他面板（如打字时脏检测弹出操作区）；
            用户手动点芯片始终为 True，保持手风琴体验。
        """
        collapsed = bool(collapsed)
        if collapsed == self._collapsed:
            return
        # 一次性标志：供 ChipBar._on_section_changed 读取
        self._no_accordion_close = not accordion_close and not collapsed
        if animate:
            self._animated_set_collapsed(collapsed)
            return
        # 即时切换：若动画仍在跑（如精简模式/自动展开与点击动画并发），
        # 必须先终止旧动画，否则旧动画结束回调会把内容显示状态写反
        self._kill_anim()
        self._collapsed = collapsed
        self._content.setVisible(not collapsed)
        self.collapsedChanged.emit(collapsed)
        self._notify_window(not collapsed)

    def _notify_window(self, expanding: bool, animate_window: bool = False) -> None:
        """通知顶层窗口弹性适应（折叠→收缩 / 展开→放大）。"""
        win = self.window()
        if win is not None and win.isVisible() and hasattr(win, "fit_to_content"):
            if animate_window:
                win.fit_to_content(expanding, animate_window=True)
            else:
                QTimer.singleShot(0, lambda: win.fit_to_content(expanding))

    def _animated_set_collapsed(self, collapsed: bool) -> None:
        """果冻式折叠/展开：展开轻柔回弹（允吸感），折叠平滑吸入；窗口弹性跟随。"""
        content = self._content
        anim = self._anim
        # 上一次动画未结束：定格当前高度（stop 不触发 finished），从该处续动
        if anim.state() == QAbstractAnimation.State.Running:
            anim.stop()
        cur_max = max(0, content.maximumHeight())
        if collapsed:
            # 折叠：从当前可见高度收到 0（完全展开时取实际高度）
            if cur_max >= 16777215:
                cur_max = content.height() or content.sizeHint().height()
        else:
            # 展开：完全折叠态（不可见）从 0 弹到理想高度；
            # 动画中途反转则从当前已弹出高度续动
            if cur_max >= 16777215:
                cur_max = 0

        self._collapsed = collapsed
        self.collapsedChanged.emit(collapsed)

        content.setVisible(True)
        content.setMaximumHeight(cur_max)
        anim.setStartValue(cur_max)
        if collapsed:
            anim.setEndValue(0)
            anim.setDuration(260)
            # 吸入感：平滑加速收起
            anim.setEasingCurve(QEasingCurve.Type.InQuart)
        else:
            target_h = content.sizeHint().height()
            if target_h <= 0:
                # 尺寸提示异常（控件尚未布局完成）：放弃动画走即时展开，避免空转
                content.setMaximumHeight(16777215)
                content.setVisible(True)
                self._notify_window(True)
                return
            anim.setEndValue(target_h)
            anim.setDuration(420)
            # 轻微 OutBack 回弹（过冲 1.05，柔和的果冻/允吸感）
            soft_curve = QEasingCurve(QEasingCurve.Type.OutBack)
            try:
                soft_curve.setOvershoot(1.05)
            except Exception:
                pass
            anim.setEasingCurve(soft_curve)
            # 展开时窗口同步放大，给内容腾出弹跳空间
            self._notify_window(True, animate_window=True)
        self._anim_target_collapsed = collapsed
        anim.start()

    def _on_anim_finished(self) -> None:
        collapsed = self._anim_target_collapsed
        if collapsed:
            self._content.setVisible(False)
        self._content.setMaximumHeight(16777215)  # 解除高度限制
        # 动画结束后窗口弹性跟随（折叠收缩 / 展开放大补齐）
        self._notify_window(not collapsed, animate_window=True)

    def toggle(self) -> None:
        self.set_collapsed(not self._collapsed, animate=True)

    def setEnabled(self, enabled: bool) -> None:
        """只禁用/启用内容区，芯片按钮始终可点（面板仍可展开查看）。"""
        self._content.setEnabled(enabled)


class FlowLayout(QLayout):
    """自动换行流式布局：窄宽度下芯片自动竖向换行（空栏收缩成芯片导轨用）。"""

    def __init__(self, parent=None, margin=0, h_spacing=6, v_spacing=4):
        super().__init__(parent)
        if parent is not None:
            self.setContentsMargins(margin, margin, margin, margin)
        self._items: List = []
        self._h_space = h_spacing
        self._v_space = v_spacing

    def __del__(self):
        while self.count():
            self.takeAt(0)

    def addItem(self, item):
        self._items.append(item)

    def count(self):
        return len(self._items)

    def itemAt(self, index):
        if 0 <= index < len(self._items):
            return self._items[index]
        return None

    def takeAt(self, index):
        if 0 <= index < len(self._items):
            return self._items.pop(index)
        return None

    def expandingDirections(self):
        return Qt.Orientation(0)

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        return self._do_layout(QRect(0, 0, width, 0), True)

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._do_layout(rect, False)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        m = self.contentsMargins()
        size += QSize(m.left() + m.right(), m.top() + m.bottom())
        return size

    def _spacing(self, pm):
        opt = None
        widget = self.parentWidget()
        if widget is not None:
            opt = widget.style()
        result = opt.pixelMetric(pm, None, widget) if opt is not None else 0
        return max(result, 0)

    def _h_spacing(self):
        return self._h_space if self._h_space >= 0 else self._spacing(
            QStyle.PixelMetric.PM_LayoutHorizontalSpacing)

    def _v_spacing(self):
        return self._v_space if self._v_space >= 0 else self._spacing(
            QStyle.PixelMetric.PM_LayoutVerticalSpacing)

    def _do_layout(self, rect, test_only):
        m = self.contentsMargins()
        effective = rect.adjusted(m.left(), m.top(), -m.right(), -m.bottom())
        x, y, line_height = effective.x(), effective.y(), 0
        for item in self._items:
            wsize = item.sizeHint()
            next_x = x + wsize.width() + self._h_spacing()
            if next_x - self._h_spacing() > effective.right() and line_height > 0:
                x = effective.x()
                y = y + line_height + self._v_spacing()
                next_x = x + wsize.width() + self._h_spacing()
                line_height = 0
            if not test_only:
                item.setGeometry(QRect(QPoint(x, y), wsize))
            x = next_x
            line_height = max(line_height, wsize.height())
        return y + line_height - rect.y() + m.bottom()


class ChipBar(QWidget):
    """一排胶囊芯片按钮：点击果冻展开/折叠对应 ChipSection。

    - exclusive=False（默认，数据发送页）：多个面板可同时展开；
    - exclusive=True（定时任务页）：手风琴，展开一个自动收起同排其他面板。
    """

    def __init__(self, exclusive: bool = False, parent=None):
        super().__init__(parent)
        self._exclusive = exclusive
        self.sections: List[ChipSection] = []
        self._syncing = False
        # 流式布局：宽度够时芯片一横排，栏收缩后芯片自动换行成"导轨"
        self._bar_layout = FlowLayout(self, margin=0, h_spacing=6, v_spacing=4)
        # 宽度随父栏收缩，高度按宽度换行自适应
        self.setMinimumWidth(108)

    def set_exclusive(self, exclusive: bool) -> None:
        self._exclusive = bool(exclusive)

    def add_section(self, title: str, collapsed: bool = False) -> ChipSection:
        """创建芯片按钮+面板并登记；面板需由调用方 addWidget 到同列布局。"""
        section = ChipSection(collapsed=collapsed)
        section._chip_bar = self
        section._title_text = title

        btn = QPushButton(title)
        btn.setCheckable(True)
        btn.setChecked(not collapsed)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet(_chip_css())
        btn.clicked.connect(
            lambda _checked=False, s=section:
                s.set_collapsed(not s.is_collapsed(), animate=True)
        )
        self._bar_layout.addWidget(btn)
        section._chip = btn
        section.collapsedChanged.connect(
            lambda c, s=section: self._on_section_changed(s, c)
        )
        self.sections.append(section)
        return section

    def _on_section_changed(self, section: ChipSection, collapsed: bool) -> None:
        # 同步芯片选中态（用户点击 / 程序自动折叠都一致）
        if section._chip is not None:
            section._chip.setChecked(not collapsed)
        # 程序自动展开且声明不挤掉其他面板（如打字中脏检测弹操作区）：跳过手风琴
        no_close = getattr(section, "_no_accordion_close", False)
        section._no_accordion_close = False
        # 手风琴：展开一个，自动收起同排其他面板
        if (self._exclusive and not collapsed and not no_close
                and not self._syncing):
            self._syncing = True
            try:
                for other in self.sections:
                    if other is not section and not other.is_collapsed():
                        other.set_collapsed(True, animate=True)
            finally:
                self._syncing = False

    def changeEvent(self, event):
        """系统/应用明暗主题切换时重刷所有芯片配色。"""
        super().changeEvent(event)
        if event.type() == QEvent.Type.PaletteChange:
            css = _chip_css()
            for s in self.sections:
                if s._chip is not None:
                    s._chip.setStyleSheet(css)


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
    progress = pyqtSignal(int, int)  # current, total
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


class WpsFileListWorker(QThread):
    """获取在线表格列表，供“浏览在线文档”对话框使用。"""

    def __init__(self, store):
        super().__init__()
        self.store = store
        self.signals = WorkerSignals()

    def run(self):
        try:
            client = WpsCloudClient(
                self.store, log_fn=lambda m: self.signals.log.emit(str(m))
            )
            raw = client.list_spreadsheets()
            items = []
            for f in raw:
                items.append((
                    f.get("name") or f.get("id") or "",
                    f.get("id", ""),
                    f.get("group_id", ""),
                    f.get("mtime", 0),
                ))
            self.signals.result.emit({"files": items})
        except Exception as e:
            self.signals.error.emit(str(e))


class CloudExcelReadWorker(QThread):
    """读取金山文档在线表格：

    下载 xlsx 到 %TEMP% → 一次性解析全部 sheet → finally 立即删除临时文件
    （读完即清缓存，切换 sheet 使用内存缓存，不重复下载）。
    """

    def __init__(self, file_token, file_name="", sheet_name=None):
        super().__init__()
        self.file_token = file_token
        self.file_name = file_name or file_token
        self.sheet_name = sheet_name
        # 兼容现有回调中对 worker.file_path 的引用
        self.file_path = f"wps-cloud://{file_token}"
        self.signals = WorkerSignals()
        self.stopped_event = threading.Event()

    def stop(self):
        self.stopped_event.set()

    def run(self):
        temp_path = None
        try:
            store = WpsOAuthStore()
            client = WpsCloudClient(
                store, log_fn=lambda m: self.signals.log.emit(str(m))
            )
            temp_path = new_temp_xlsx_path()
            client.download_to(self.file_token, temp_path)
            if self.stopped_event.is_set():
                return

            all_sheets = read_all_sheets(
                temp_path,
                log_fn=lambda m: self.signals.log.emit(str(m)),
            )
            if not all_sheets:
                raise Exception("在线表格为空")

            sheet_names = list(all_sheets.keys())
            current_sheet = (
                self.sheet_name
                if self.sheet_name and self.sheet_name in all_sheets
                else sheet_names[0]
            )
            values = all_sheets[current_sheet]
            if not values:
                raise Exception(f"工作表 {current_sheet} 为空")

            headers = TableProcessor.derive_headers(values[0])
            self.signals.log.emit(f"在线文档读取完成：{self.file_name} / {current_sheet}")
            self.signals.result.emit({
                "sheet_names": sheet_names,
                "current_sheet": current_sheet,
                "headers": headers,
                "data": values,
                "all_sheets_data": all_sheets,
                "source": "wps_cloud",
                "cloud_file_id": self.file_token,
                "cloud_file_name": self.file_name,
            })
        except Exception as e:
            import traceback
            self.signals.log.emit(f"✗ 在线文档读取失败: {e}")
            self.signals.log.emit(traceback.format_exc()[:300])
            self.signals.error.emit(str(e))
        finally:
            # 关键：读完立即删除临时文件，不落地缓存
            cleanup_temp_file(temp_path)


class SendWorker(QThread):
    def __init__(
        self,
        tasks,
        send_interval=2,
        chat_delay=0.8,
        send_order=None,
        send_mode=None,
        minimize_after=True,
        attachment=""
    ):
        super().__init__()
        self.tasks = tasks
        self.send_interval = send_interval
        self.chat_delay = chat_delay
        self.minimize_after = bool(minimize_after)
        self.attachment = str(attachment or "").strip()
        # 发送顺序：显式 send_order 优先；兼容旧 send_mode 参数
        if send_order:
            order = list(send_order)
        elif send_mode in ("image",):
            order = ["image"]
        elif send_mode == "image_text":
            order = ["image", "text"]
        else:
            order = ["text"]
        # 去重保序，只保留合法类型
        self.send_order = []
        for k in order:
            if k in ("text", "image", "custom", "attachment") and k not in self.send_order:
                self.send_order.append(k)
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
        """按用户配置的发送顺序生成步骤；内容为空的步骤跳过。"""
        steps = []
        for kind in self.send_order:
            if kind == "image":
                steps.append({"type": "image", "index": 0})
            elif kind == "text":
                steps.append({"type": "text", "index": 0})
            elif kind == "custom" and custom_msg:
                steps.append({"type": "custom", "index": 0})
            elif kind == "attachment" and self.attachment:
                steps.append({"type": "attachment", "index": 0})
        return steps

    def run(self):
        from modules.send_executor import run_send_tasks

        total_count = len(self.tasks)
        sender = None
        result_emitted = False
        acquired = False  # 是否已获得全局发送批次锁
        # 预初始化：若在赋值前抛异常，finally 还原时不能引用未定义变量
        _orig_sender_log = None

        # 附加文件预检：配置了附件但文件不存在时提示并按无附件继续。
        if self.attachment and not os.path.exists(self.attachment):
            self.signals.log.emit(
                f"\u26a0 附加文件不存在，本次发送不带附件: {self.attachment}"
            )
            self.attachment = ""

        try:
            # 全局发送批次锁：与其他发送任务（定时/监控）排队串行
            acquired = WeChatSender.acquire_batch(
                log_fn=self.signals.log.emit,
                should_stop=self.stopped_event.is_set,
            )
            if not acquired:
                self.signals.log.emit("❌ 发送已停止，取消排队")
                self.signals.result.emit((0, total_count, total_count, list(self.tasks)))
                result_emitted = True
                return
            self.tasks = [self._normalize_task(t) for t in self.tasks]
            self.signals.log.emit("初始化微信客户端...")
            sender = WeChatSender.shared_instance()
            sender.cleanup_temp_images()
            if not sender.initialize():
                # 单例失败不影响下次，清空让后续尝试重新 new WeChat
                WeChatSender.reset_shared_instance()
                # 只走 result 单通道（失败数=总数，支持一键重试）
                self.signals.log.emit("❌ 微信未登录或未打开，本次发送全部失败")
                self.signals.result.emit((0, total_count, total_count, list(self.tasks)))
                result_emitted = True
                return
            self.signals.log.emit("微信客户端初始化成功")
            # 保存原始 log 方法，在 finally 中还原，避免共享单例的 log
            # 指向已销毁 QThread 的信号
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

            success_count, failed_tasks = run_send_tasks(
                self.tasks,
                sender,
                send_order=self.send_order,
                attachment=self.attachment,
                send_interval=self.send_interval,
                chat_delay=self.chat_delay,
                log=self.signals.log.emit,
                progress=lambda c, t: self.signals.progress.emit(c, t),
                stop_event=self.stopped_event,
                pause_event=self.paused_event,
            )

            failed_count = len(failed_tasks)
            if failed_tasks:
                self.signals.log.emit(f"\n--- 发送失败列表 ({failed_count}人) ---")
                for failed_task in failed_tasks:
                    self.signals.log.emit(
                        f"❌ {failed_task['name']} → {failed_task['recipient']}"
                    )

            self.signals.log.emit(
                f"\n发送完成！成功: {success_count}, "
                f"失败: {failed_count}, 总计: {total_count}"
            )
            self.signals.result.emit(
                (success_count, failed_count, total_count, failed_tasks)
            )
            result_emitted = True

        except Exception as e:
            import traceback
            self.signals.log.emit(f"❌ 发送线程异常: {str(e)}")
            self.signals.log.emit(f"详细错误: {traceback.format_exc()[:300]}")
            # 统一只走 result 单通道（错误详情已在上面日志给出）
            if not result_emitted:
                self.signals.result.emit(
                    (0, total_count, total_count, list(self.tasks))
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
                if _orig_sender_log is not None:
                    try:
                        sender.log = _orig_sender_log
                    except Exception:
                        pass
                try:
                    sender.cleanup_temp_images()
                except Exception:
                    pass
            # 释放全局发送批次锁（排在最后，保证窗口/清理完成后才轮到下一批次）
            if acquired:
                try:
                    WeChatSender.release_batch()
                except Exception:
                    pass

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
    3) 同时写入 GUI 面板和 app.log（logging）：
       定时任务无人值守执行，仅写面板时一旦任务卡死/异常，事后无任何
       文件日志可查（2026-09-28 孙中枢弹窗卡死事件即因此无日志定位）。
    """
    __slots__ = ("_cb", "_lg")

    def __init__(self, cb):
        self._cb = cb
        self._lg = logging.getLogger("wechat_sender").info

    def log(self, message):
        try:
            self._lg(str(message))
        except Exception:
            pass
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
    progress = pyqtSignal(int, int)  # current, total：用于任务栏/标题进度（主线程连接）

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
        attach_order="after",
        send_order=None,
    ):
        super().__init__()
        self.recipients = list(recipients or [])
        self.message = message or ""
        self.chat_delay = max(0.0, min(10.0, float(chat_delay)))
        self.send_interval = max(0.0, min(30.0, float(send_interval)))
        self.default_city = str(default_city or "")
        self.minimize_after = bool(minimize_after)
        self.attachment = str(attachment or "").strip()
        # 发送顺序：显式 send_order 优先；否则按 attach_order 兼容旧调用
        if isinstance(send_order, list) and send_order:
            order = [k for k in send_order if k in ("message", "attachment")]
            if "message" not in order:
                order.append("message")
            self.send_order = order
        else:
            self.send_order = (
                ["attachment", "message"] if str(attach_order) == "before"
                else ["message", "attachment"]
            )
        self._log_cb = log_callback
        self.stopped_event = threading.Event()
        self._sender = None
        # 是否已获得全局发送批次锁（供调度线程区分"排队中"与"已开始发送"）
        self._batch_acquired = False

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
            # 全局发送批次锁：与其他发送任务（手动/监控）排队串行
            self._batch_acquired = WeChatSender.acquire_batch(
                log_fn=self.log,
                should_stop=self.stopped_event.is_set,
            )
            if not self._batch_acquired:
                self.log("[定时] ❌ 已停止，取消排队")
                self.finished_with_result.emit(0, len(self.recipients), list(self.recipients))
                return
            self._sender = WeChatSender.shared_instance()
            # 保存原始 log 方法，在结束时还原，避免共享单例的 log 指向已销毁的 QThread
            _orig_sender_log = self._sender.log
            if self._log_cb:
                self._sender.log = _SenderLogCb(self._log_cb).log
            if not self._sender.initialize():
                WeChatSender.reset_shared_instance()
                self.log("[定时] 微信初始化失败，本次任务失败")
                if self._batch_acquired:
                    try:
                        WeChatSender.release_batch()
                    except Exception:
                        pass
                    self._batch_acquired = False
                self.finished_with_result.emit(0, len(self.recipients), list(self.recipients))
                return
            self.log("[定时] 微信客户端初始化成功")
        except Exception as exc:
            self.log(f"[定时] 初始化异常: {exc}")
            WeChatSender.reset_shared_instance()
            if self._batch_acquired:
                try:
                    WeChatSender.release_batch()
                except Exception:
                    pass
                self._batch_acquired = False
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
        current_idx = 0
        try:
            for idx, recipient in enumerate(self.recipients):
                current_idx = idx
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
                self.progress.emit(idx + 1, total)

                text_ok = False
                # 每个收件人只在第一个发送项前切换一次聊天窗口；后续
                # 文字/附件直接在当前窗口用 fast_mode 发送，不再重复
                # "搜索-切换-确认"（与数据发送 Tab 逻辑一致）。
                chat_opened = False
                # v1.5.0 批处理合并：文字 + 附件合并为一次 send_text_and_files
                # （一次粘贴 + 一次 Enter），绝不逐条发；仅附件时单独发文件。
                # 接收人成败以消息文字为准，附件失败不计入失败。
                kinds = [k for k in self.send_order if k in ("message", "attachment")]
                has_msg = "message" in kinds
                has_attach = "attachment" in kinds and attachment_ok
                # 发送顺序联动粘贴顺序：message 排在 attachment 前 → 先粘文字再粘附件
                text_first = bool(kinds) and kinds[0] == "message"
                if has_msg or has_attach:
                    if self.stopped_event.is_set():
                        break
                    if not self._sender.open_chat(
                        recipient,
                        chat_delay=self.chat_delay,
                        stop_event=self.stopped_event,
                    ):
                        self.log(f"[定时] ✗ 打开聊天窗口失败: {recipient}")
                        chat_opened = False
                    else:
                        chat_opened = True
                if chat_opened:
                    if has_msg and has_attach:
                        self.log(f"[定时] 合并发送文字+附件给 {recipient}")
                        text_ok = self._sender.send_text_and_files(
                            text=self.message,
                            file_paths=[self.attachment],
                            recipient=recipient,
                            chat_delay=self.chat_delay,
                            fast_mode=True,
                            stop_event=self.stopped_event,
                            text_first=text_first,
                        )
                    elif has_msg:
                        text_ok = self._sender.send_message(
                            content=self.message,
                            recipient=recipient,
                            chat_delay=self.chat_delay,
                            fast_mode=True,
                            stop_event=self.stopped_event,
                        )
                    elif has_attach:
                        if not self._sender.send_file(
                            self.attachment,
                            recipient,
                            chat_delay=self.chat_delay,
                            fast_mode=True,
                            stop_event=self.stopped_event,
                        ):
                            self.log(f"[定时] ⚠ 附加文件发送失败: {recipient}")
                            text_ok = False
                        else:
                            text_ok = True
                ok = text_ok
                if ok:
                    success += 1
                else:
                    failed_recipients.append(recipient)
                    self.log(f"[定时] ✗ 发送失败: {recipient}")
        except Exception as exc:
            # 循环内任何未预期异常都不能让结束信号丢失，
            # 否则主窗口标题会永远停在“发送中…”
            try:
                self.log(f"[定时] ⚠ 发送循环异常中断: {exc}")
            except Exception:
                pass
            # 当前及之后的接收人都未完成，补入失败列表（去重）以便用户重试
            existing = set(failed_recipients)
            for r in self.recipients[current_idx:]:
                r = str(r).strip()
                if r and r not in existing:
                    failed_recipients.append(r)
                    existing.add(r)
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
            # 释放全局发送批次锁（排在最后，保证窗口/清理完成后才轮到下一批次）
            if self._batch_acquired:
                try:
                    WeChatSender.release_batch()
                except Exception:
                    pass
                self._batch_acquired = False
            # 关键：无论正常结束、手动停止还是异常中断，都必须发出结束信号，
            # 主线程据此恢复标题/清除任务栏进度
            self.finished_with_result.emit(
                success, len(failed_recipients), failed_recipients
            )


class ProfileChainWorker(QThread):
    """按顺序执行多个配置文件。

    每个配置：读本地/云端表格（云端临时文件读完即删）→ 按配置筛选人员
    → 发送表格文字/图片/自定义消息/附件；一个配置完成后再执行下一个。
    微信客户端全程只初始化一次。
    """

    finished_with_result = pyqtSignal(int, int, list)  # success, failed, failed_names
    progress = pyqtSignal(int, int)  # 第几个配置 / 配置总数
    send_progress = pyqtSignal(int, int, int, int)  # 链路当前, 链路总数, 发送当前, 发送总数

    def __init__(self, profile_paths, log_callback=None, minimize_after=True):
        super().__init__()
        self.profile_paths = [str(p) for p in (profile_paths or [])]
        self.minimize_after = bool(minimize_after)
        self._log_cb = log_callback
        self.signals = WorkerSignals()
        self.stopped_event = threading.Event()
        self._sender = None
        # 是否已获得全局发送批次锁（供调度线程区分"排队中"与"已开始发送"）
        self._batch_acquired = False

    def stop(self):
        self.stopped_event.set()

    def log(self, msg):
        try:
            self.signals.log.emit(str(msg))
        except Exception:
            pass
        if self._log_cb:
            try:
                self._log_cb(msg)
            except Exception:
                pass

    def run(self):
        from modules.config_manager import ConfigManager, ConfigError
        from modules.profile_runner import prepare_profile, ProfilePrepareError
        from modules.send_executor import run_send_tasks

        total = len(self.profile_paths)
        success_total = 0
        failed_names = []
        sender = None
        _orig_sender_log = None
        # 结束信号全函数只允许发一次（early return 与 finally 竞态保护）
        _finished_emitted = False

        def _emit_finished():
            nonlocal _finished_emitted
            if _finished_emitted:
                return
            _finished_emitted = True
            try:
                self.finished_with_result.emit(
                    success_total, len(failed_names), failed_names
                )
            except Exception:
                pass

        if total == 0:
            self.log("[配置链] 没有要执行的配置文件")
            _emit_finished()
            return

        try:
            if is_workstation_locked():
                self.log(
                    "[配置链] ⚠ 电脑处于锁定状态且无法自动解锁，本次发送很可能失败；"
                    "请保持电脑解锁或开启'执行期间防自动锁定'"
                )
        except Exception:
            pass

        try:
            # 全局发送批次锁：与其他发送任务（手动/定时/监控）排队串行
            self._batch_acquired = WeChatSender.acquire_batch(
                log_fn=self.log,
                should_stop=self.stopped_event.is_set,
            )
            if not self._batch_acquired:
                self.log("[配置链] ❌ 已停止，取消排队")
                _emit_finished()
                return
            self.log("[配置链] 初始化微信客户端...")
            sender = WeChatSender.shared_instance()
            _orig_sender_log = sender.log
            if self._log_cb:
                sender.log = _SenderLogCb(self._log_cb).log
            try:
                init_ok = bool(sender.initialize())
            except Exception as exc:
                self.log(f"[配置链] ❌ 微信初始化异常：{exc}")
                init_ok = False
            if not init_ok:
                WeChatSender.reset_shared_instance()
                self.log("[配置链] ❌ 微信初始化失败，本次任务失败")
                failed_names.append("微信初始化失败")
                _emit_finished()
                return
            sender.cleanup_temp_images()
            self.log("[配置链] 微信客户端初始化成功")

            mgr = ConfigManager()
            for idx, path in enumerate(self.profile_paths):
                if self.stopped_event.is_set():
                    self.log("[配置链] 已手动停止，后续配置不再执行")
                    remaining = [
                        os.path.splitext(os.path.basename(p))[0]
                        for p in self.profile_paths[idx:]
                    ]
                    failed_names.extend(remaining)
                    break
                self.progress.emit(idx + 1, total)
                name = os.path.splitext(os.path.basename(path))[0]
                self.log(f"[配置链] ({idx + 1}/{total}) 开始执行配置「{name}」")
                try:
                    profile = mgr.load_profile(path)
                    prep = prepare_profile(
                        profile, log_fn=lambda m: self.log(str(m))
                    )
                    tasks = prep["tasks"]
                    self.log(
                        f"[配置链] 「{name}」准备完成："
                        f"{prep['sheet']} 共 {len(tasks)} 人，"
                        f"顺序 {'/'.join(prep['send_order']) or '(空)'}"
                    )
                    ok, failed = run_send_tasks(
                        tasks,
                        sender,
                        send_order=prep["send_order"],
                        attachment=prep["attachment"],
                        send_interval=prep["send_interval"],
                        chat_delay=prep["chat_delay"],
                        log=self.log,
                        progress=lambda c, t: self.send_progress.emit(
                            idx + 1, total, c, t),
                        stop_event=self.stopped_event,
                        label=f"{idx + 1}/{total}",
                    )
                    success_total += ok
                    failed_names.extend(t["name"] for t in failed)
                    self.log(
                        f"[配置链] 「{name}」完成：成功 {ok}，失败 {len(failed)}"
                    )
                except (ConfigError, ProfilePrepareError) as exc:
                    self.log(f"[配置链] ❌ 配置「{name}」无法执行：{exc}")
                    failed_names.append(name)
                except Exception as exc:
                    self.log(f"[配置链] ❌ 配置「{name}」执行异常：{exc}")
                    failed_names.append(name)
        except Exception as exc:
            # 初始化/外层任何意外异常都不能丢结束信号，否则标题永久“发送中…”
            self.log(f"[配置链] ❌ 任务异常中断：{exc}")
        finally:
            if sender is not None:
                if self.minimize_after:
                    try:
                        sender.minimize_window()
                    except Exception:
                        pass
                if _orig_sender_log is not None:
                    try:
                        sender.log = _orig_sender_log
                    except Exception:
                        pass
            # 释放全局发送批次锁
            if self._batch_acquired:
                try:
                    WeChatSender.release_batch()
                except Exception:
                    pass
            self.log(
                f"[配置链] 全部结束：{total} 个配置，"
                f"累计成功 {success_total}，失败 {len(failed_names)}"
            )
            _emit_finished()


class WpsCredentialDialog(QDialog):
    """金山文档 Cookie 凭证（wps_sid）设置对话框。"""
    credentials_changed = pyqtSignal()

    def __init__(self, store: WpsOAuthStore, parent=None):
        super().__init__(parent)
        self.store = store
        self.setWindowTitle("云文档 Cookie 凭证设置（金山文档）")
        self.setMinimumWidth(560)

        layout = QVBoxLayout(self)

        tip = QLabel(
            "使用浏览器登录态（wps_sid）读取金山文档，无需申请开放平台应用：\n"
            "1. 用 Edge/Chrome 打开 https://www.kdocs.cn 并登录你的账号；\n"
            "2. 按 F12 → 应用/Application → Cookie → https://www.kdocs.cn；\n"
            "3. 找到 wps_sid，复制它的值粘贴到下面保存（仅保存在本机）。\n"
            "Cookie 失效后（读取报登录失效），重新复制一次即可。"
        )
        tip.setWordWrap(True)
        tip.setStyleSheet("color: #555; font-size: 12px;")
        layout.addWidget(tip)

        row = QHBoxLayout()
        row.addWidget(QLabel("wps_sid:"))
        self.sid_edit = QLineEdit()
        self.sid_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.sid_edit.setPlaceholderText("浏览器 Cookie 中 wps_sid 的值")
        row.addWidget(self.sid_edit, 1)
        self.show_sid_btn = QPushButton("显示")
        self.show_sid_btn.setCheckable(True)
        self.show_sid_btn.toggled.connect(self._toggle_sid_visible)
        row.addWidget(self.show_sid_btn)
        layout.addLayout(row)

        btn_row = QHBoxLayout()
        self.save_cred_btn = QPushButton("💾 保存凭证")
        self.save_cred_btn.clicked.connect(self._save_credentials)
        btn_row.addWidget(self.save_cred_btn)
        self.verify_btn = QPushButton("🔍 验证登录态")
        self.verify_btn.setStyleSheet(
            "background-color: #009688; color: white; padding: 6px 12px;"
        )
        self.verify_btn.clicked.connect(self._verify_cookie)
        btn_row.addWidget(self.verify_btn)
        self.reauth_btn = QPushButton("清除凭证")
        self.reauth_btn.setToolTip("删除本机保存的 wps_sid")
        self.reauth_btn.clicked.connect(self._clear_tokens)
        btn_row.addWidget(self.reauth_btn)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)

        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet("font-size: 12px;")
        layout.addWidget(self.status_label)

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setPlaceholderText("验证/读取日志…")
        self.log_edit.setMaximumHeight(120)
        layout.addWidget(self.log_edit)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Close
        )
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

        self.sid_edit.setText(self.store.wps_sid)
        self._refresh_status()

    def _toggle_sid_visible(self, checked):
        self.sid_edit.setEchoMode(
            QLineEdit.EchoMode.Normal
            if checked
            else QLineEdit.EchoMode.Password
        )
        self.show_sid_btn.setText("隐藏" if checked else "显示")

    def _log(self, msg):
        self.log_edit.append(str(msg))

    def _save_credentials(self):
        sid = self.sid_edit.text().strip()
        if not sid:
            QMessageBox.warning(self, "提示", "请先粘贴 wps_sid 的值")
            return
        self.store.data["wps_sid"] = sid
        try:
            self.store.save()
        except OSError as exc:
            QMessageBox.warning(self, "保存失败", f"凭证写入失败：{exc}")
            return
        self._log("Cookie 凭证已保存到本机配置文件")
        self.credentials_changed.emit()
        self._refresh_status()
        QMessageBox.information(self, "已保存", "wps_sid 已保存到本机。")

    def _verify_cookie(self):
        sid = self.sid_edit.text().strip()
        if not sid:
            QMessageBox.warning(self, "提示", "请先粘贴 wps_sid 的值")
            return
        self.store.data["wps_sid"] = sid
        try:
            self.store.save()
        except OSError:
            pass
        self.verify_btn.setEnabled(False)
        self.status_label.setText("⏳ 正在验证登录态…")
        self.status_label.setStyleSheet("color: #EF6C00; font-size: 12px;")
        self._log("调用金山文档接口验证 wps_sid…")
        worker = WpsFileListWorker(self.store)
        self.verify_worker = worker
        worker.signals.log.connect(self._log)
        worker.signals.result.connect(self._on_verify_result)
        worker.signals.error.connect(self._on_verify_error)
        worker.start()

    def _on_verify_result(self, result):
        self.verify_btn.setEnabled(True)
        files = (result or {}).get("files") or []
        self.status_label.setText(
            f"✅ 登录态有效，识别到 {len(files)} 个在线表格"
        )
        self.status_label.setStyleSheet("color: #2E7D32; font-size: 12px;")
        self.credentials_changed.emit()

    def _on_verify_error(self, error):
        self.verify_btn.setEnabled(True)
        self.status_label.setText(f"❌ 验证失败：{error}")
        self.status_label.setStyleSheet("color: #C62828; font-size: 12px;")
        self._log(f"验证失败：{error}")

    def _clear_tokens(self):
        self.store.clear_tokens()
        self.sid_edit.clear()
        self._log("已清除本机 Cookie 凭证")
        self.credentials_changed.emit()
        self._refresh_status()

    def _refresh_status(self):
        if not self.store.is_configured():
            self.status_label.setText("状态：未配置（请粘贴浏览器里的 wps_sid）")
            self.status_label.setStyleSheet("color: #C62828; font-size: 12px;")
        else:
            self.status_label.setText("状态：已保存 wps_sid（失效时重新复制即可）")
            self.status_label.setStyleSheet("color: #2E7D32; font-size: 12px;")

    def reject(self):
        # 若验证请求仍在进行，等线程收尾，避免 QThread 运行中被销毁
        worker = getattr(self, "verify_worker", None)
        if worker is not None and worker.isRunning():
            worker.wait(3000)
        super().reject()


class CloudFilePickerDialog(QDialog):
    """从账号下在线表格列表中选择一个云文档。"""

    def __init__(self, files, parent=None):
        super().__init__(parent)
        self.selected = None
        self.setWindowTitle("选择在线表格")
        self.resize(560, 460)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("账号下的在线表格（双击选择）："))

        from datetime import datetime
        self.list_widget = QListWidget()
        for name, file_id, group_id, mtime in files:
            when = (
                datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
                if mtime
                else ""
            )
            item = QListWidgetItem(f"{name}    {when}")
            item.setData(
                Qt.ItemDataRole.UserRole, (name, file_id, group_id)
            )
            self.list_widget.addItem(item)
        self.list_widget.itemDoubleClicked.connect(self._accept_item)
        layout.addWidget(self.list_widget, 1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _accept_item(self, item):
        self.selected = item.data(Qt.ItemDataRole.UserRole)
        self.accept()

    def _on_ok(self):
        item = self.list_widget.currentItem()
        if item:
            self.selected = item.data(Qt.ItemDataRole.UserRole)
        self.accept()


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
        self._current_header_row = 1
        # 数据来源：local 本地文件 / wps_cloud 金山文档在线表格
        self.data_source = "local"
        self.wps_store = WpsOAuthStore()
        self.cloud_file_id = ""
        self.cloud_file_name = ""
        # 在线表格一次下载全部 sheet 后缓存在内存，切 sheet 不重复下载
        self._cloud_sheets_cache = {}
        self.cloud_worker = None
        self._cloud_list_worker = None
        self.excel_worker = None
        self.worker = None
        # 由 MainWindow 注入：worker 创建后回调，用于连接任务栏进度等全局信号
        self.worker_created_cb = None
        # 由 MainWindow 注入：完成/错误弹窗被用户点掉后回调，用于恢复任务栏与标题
        self.progress_dismissed_cb = None
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

        # 左栏芯片排：配置管理 / 文档设置 / 筛选条件（可多个同时展开）
        left_bar = ChipBar(exclusive=False)
        left_layout.addWidget(left_bar)

        config_group = left_bar.add_section("⚙ 配置管理", collapsed=True)
        config_layout = config_group.contentLayout()

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

        self.edit_config_btn = QPushButton("编辑配置")
        self.edit_config_btn.setToolTip(
            "把配置加载到界面进行修改，但不触发自动发送（适合编辑"
            "勾选了「加载后自动发送」的配置）；改完点「保存配置」即可"
        )
        config_btn_layout.addWidget(self.edit_config_btn)
        config_layout.addLayout(config_btn_layout)

        config_layout.addWidget(QLabel("最近配置（双击加载，右键可编辑）:"))
        self.recent_config_list = QListWidget()
        self.recent_config_list.setMinimumHeight(50)
        self.recent_config_list.setMaximumHeight(72)
        # 右键菜单：加载 / 编辑（不自动发送）
        self.recent_config_list.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu)
        config_layout.addWidget(self.recent_config_list)

        left_layout.addWidget(config_group)

        url_group = left_bar.add_section("📄 文档设置")
        url_layout = url_group.contentLayout()

        # 数据来源切换：本地文件 / 金山文档在线表格
        source_layout = QHBoxLayout()
        source_layout.addWidget(QLabel("数据来源:"))
        self.local_source_radio = QRadioButton("本地文件")
        self.cloud_source_radio = QRadioButton("☁ 金山在线文档")
        self.local_source_radio.setChecked(True)
        self.source_button_group = QButtonGroup(self)
        self.source_button_group.addButton(self.local_source_radio)
        self.source_button_group.addButton(self.cloud_source_radio)
        source_layout.addWidget(self.local_source_radio)
        source_layout.addWidget(self.cloud_source_radio)
        source_layout.addStretch(1)
        self.wps_settings_btn = QPushButton("Cookie 凭证设置")
        self.wps_settings_btn.setStyleSheet(
            "background-color: #607D8B; color: white; padding: 3px 8px; font-size: 11px;"
        )
        source_layout.addWidget(self.wps_settings_btn)
        url_layout.addLayout(source_layout)

        # 本地文件行
        self.local_row_widget = QWidget()
        excel_btn_layout = QHBoxLayout(self.local_row_widget)
        excel_btn_layout.setContentsMargins(0, 0, 0, 0)
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

        url_layout.addWidget(self.local_row_widget)

        # 在线文档行（默认隐藏，切到云来源时显示）
        self.cloud_row_widget = QWidget()
        cloud_layout = QHBoxLayout(self.cloud_row_widget)
        cloud_layout.setContentsMargins(0, 0, 0, 0)
        self.cloud_file_edit = QLineEdit()
        self.cloud_file_edit.setPlaceholderText("在线文档 ID 或 kdocs.cn 链接")
        self.cloud_file_edit.setEnabled(False)
        cloud_layout.addWidget(self.cloud_file_edit, 1)
        self.cloud_browse_btn = QPushButton("浏览…")
        self.cloud_browse_btn.setStyleSheet("padding: 4px 8px; font-size: 12px;")
        self.cloud_browse_btn.setEnabled(False)
        cloud_layout.addWidget(self.cloud_browse_btn)
        self.cloud_read_btn = QPushButton("☁ 读取")
        self.cloud_read_btn.setStyleSheet(
            "background-color: #009688; color: white; padding: 4px 8px; font-size: 12px;"
        )
        self.cloud_read_btn.setEnabled(False)
        cloud_layout.addWidget(self.cloud_read_btn)
        self.cloud_row_widget.setVisible(False)
        url_layout.addWidget(self.cloud_row_widget)

        self.cloud_status_label = QLabel("")
        self.cloud_status_label.setWordWrap(True)
        self.cloud_status_label.setStyleSheet("color: #666; font-size: 11px;")
        self.cloud_status_label.setVisible(False)
        url_layout.addWidget(self.cloud_status_label)

        self.current_file_label = QLabel("当前文件: 未选择")
        self.current_file_label.setWordWrap(True)
        self.current_file_label.setStyleSheet("color: #666; font-size: 12px;")
        url_layout.addWidget(self.current_file_label)

        url_layout.addWidget(QLabel("Sheet名称:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setPlaceholderText("请先选择Excel文件")
        self.sheet_combo.setEnabled(False)
        url_layout.addWidget(self.sheet_combo)

        url_layout.addWidget(QLabel("表头所在行:"))
        self.header_row_spin = QSpinBox()
        self.header_row_spin.setMinimum(1)
        self.header_row_spin.setMaximum(1)
        self.header_row_spin.setValue(1)
        self.header_row_spin.setEnabled(False)
        self.header_row_spin.setToolTip(
            "分类标题（表头）所在的 Excel 行号，默认为第 1 行；\n"
            "若表头在第 2 行或更靠下，请改成对应行号后再选择列、加载数据。"
        )
        url_layout.addWidget(self.header_row_spin)

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

        filter_group = left_bar.add_section("🔍 筛选条件", collapsed=True)
        filter_layout = filter_group.contentLayout()
        
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

        # 中栏芯片排：人员列表 / 数据预览
        middle_bar = ChipBar(exclusive=False)
        middle_layout.addWidget(middle_bar)

        persons_group = middle_bar.add_section("👥 人员列表")
        persons_layout = persons_group.contentLayout()
        
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
        
        preview_group = middle_bar.add_section("👀 数据预览", collapsed=True)
        preview_layout = preview_group.contentLayout()
        
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

        # 右栏芯片排：发送设置 / 发送进度 / 日志 / 发送控制（默认只打开发送设置，
        # 发送时自动切换为只保留进度+控制）
        right_bar = ChipBar(exclusive=False)
        right_layout.addWidget(right_bar)

        send_group = right_bar.add_section("✉ 发送设置")
        send_layout = send_group.contentLayout()

        self.auto_send_check = QCheckBox(
            "⚡ 加载本配置后自动全选筛选人员并直接发送"
        )
        self.auto_send_check.setToolTip(
            "勾选后保存配置：以后每次加载该配置（含定时任务/配置链调用），\n"
            "应用筛选条件、自动全选筛选结果中的人员并立即开始发送，无需手动确认。\n"
            "请确认筛选条件准确后再开启，避免误发。"
        )
        send_layout.addWidget(self.auto_send_check)

        send_layout.addWidget(QLabel("微信接收人(手动指定):"))
        self.wechat_edit = QLineEdit()
        send_layout.addWidget(self.wechat_edit)
        
        send_mode_layout = QHBoxLayout()
        send_mode_layout.addWidget(QLabel("发送内容顺序(勾选启用，选中后点右侧按钮调整先后):"))
        send_layout.addLayout(send_mode_layout)
        order_row = QHBoxLayout()
        self.send_order_list = QListWidget()
        self.send_order_list.setMaximumHeight(92)
        # 顺序即发送先后：每项带勾选框表示是否发送该内容
        self._ORDER_ITEMS = [
            ("text", "表格文字（提取的数据）"),
            ("image", "表格图片（数据截图）"),
            ("custom", "自定义消息"),
            ("attachment", "附加文件"),
        ]
        for key, label in self._ORDER_ITEMS:
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, key)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(
                Qt.CheckState.Checked if key == "text" else Qt.CheckState.Unchecked
            )
            self.send_order_list.addItem(item)
        order_row.addWidget(self.send_order_list, 1)
        order_btns = QVBoxLayout()
        self.order_up_btn = QPushButton("⬆ 上移")
        self.order_down_btn = QPushButton("⬇ 下移")
        order_btns.addWidget(self.order_up_btn)
        order_btns.addWidget(self.order_down_btn)
        order_btns.addStretch()
        order_row.addLayout(order_btns)
        send_layout.addLayout(order_row)
        # 保留一个隐藏兼容属性：部分旧逻辑读取 send_mode_combo 时不再报错
        self.send_mode_combo = None

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

        # 联系人映射面板：把"人员列"筛选出来的值（可能是组名）映射到微信接收人
        recipient_group = right_bar.add_section("🔗 联系人映射", collapsed=True)
        recipient_layout = recipient_group.contentLayout()

        self.recipient_mapping_enabled_check = QCheckBox(
            "启用联系人映射（启用后，下方映射表会替代\"微信接收人\"逻辑）"
        )
        self.recipient_mapping_enabled_check.setToolTip(
            "勾选后：\n"
            "  • 命中映射 → 按映射展开（一对多时一个组发给多个人，每人都收完整数据）\n"
            "  • 未命中但有兜底接收人 → 用兜底\n"
            "  • 未命中且无兜底 → 回退原\"筛选列\"逻辑（wechat_column / 手动 / 筛选值本身）\n"
            "不勾选时本面板被忽略，完全走旧逻辑（向后兼容）。"
        )
        recipient_layout.addWidget(self.recipient_mapping_enabled_check)

        recipient_layout.addWidget(QLabel("兜底接收人(未命中映射时使用，可选):"))
        self.recipient_mapping_default_edit = QLineEdit()
        self.recipient_mapping_default_edit.setPlaceholderText(
            "未命中映射时的兜底接收人（可选）"
        )
        recipient_layout.addWidget(self.recipient_mapping_default_edit)

        self.recipient_mapping_table = QTableWidget(0, 2)
        self.recipient_mapping_table.setHorizontalHeaderLabels(
            ["筛选值", "接收人(多人用 / 或 ; 分隔)"]
        )
        self.recipient_mapping_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.recipient_mapping_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        self.recipient_mapping_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.recipient_mapping_table.setMinimumHeight(140)
        recipient_layout.addWidget(self.recipient_mapping_table)

        rm_btn_row1 = QHBoxLayout()
        self.rm_add_row_btn = QPushButton("➕ 添加行")
        self.rm_del_row_btn = QPushButton("➖ 删除选中行")
        self.rm_clear_btn = QPushButton("🗑 清空")
        rm_btn_row1.addWidget(self.rm_add_row_btn)
        rm_btn_row1.addWidget(self.rm_del_row_btn)
        rm_btn_row1.addWidget(self.rm_clear_btn)
        recipient_layout.addLayout(rm_btn_row1)

        rm_btn_row2 = QHBoxLayout()
        self.rm_import_btn = QPushButton("📥 导入表格")
        self.rm_open_file_btn = QPushButton("📂 打开映射表")
        self.rm_link_file_btn = QPushButton("🔗 关联文件")
        rm_btn_row2.addWidget(self.rm_import_btn)
        rm_btn_row2.addWidget(self.rm_open_file_btn)
        rm_btn_row2.addWidget(self.rm_link_file_btn)
        recipient_layout.addLayout(rm_btn_row2)

        right_layout.addWidget(recipient_group)

        # 发送预检缓存：控制"发送前验证微信联系人"结果缓存的有效期与手动清理。
        # 有效期选「永久」即不自动清理，需手动点「清除缓存」。
        precheck_group = right_bar.add_section("⚙️ 发送预检缓存", collapsed=True)
        precheck_layout = precheck_group.contentLayout()

        ttl_row = QHBoxLayout()
        ttl_row.addWidget(QLabel("缓存有效期:"))
        self.precheck_ttl_combo = QComboBox()
        self.precheck_ttl_combo.addItem("5 分钟", 300)
        self.precheck_ttl_combo.addItem("30 分钟", 1800)
        self.precheck_ttl_combo.addItem("1 小时", 3600)
        self.precheck_ttl_combo.addItem("24 小时", 86400)
        self.precheck_ttl_combo.addItem("永久（不自动清理）", 0)
        self.precheck_ttl_combo.setFixedWidth(170)
        ttl_row.addWidget(self.precheck_ttl_combo)
        precheck_layout.addLayout(ttl_row)

        self.precheck_cache_info = QLabel("预检缓存: -")
        precheck_layout.addWidget(self.precheck_cache_info)

        self.precheck_cache_clear_btn = QPushButton("🧹 清除缓存")
        precheck_layout.addWidget(self.precheck_cache_clear_btn)

        self.precheck_ttl_combo.currentIndexChanged.connect(
            self._on_precheck_ttl_changed
        )
        self.precheck_cache_clear_btn.clicked.connect(
            self._on_clear_precheck_cache
        )

        right_layout.addWidget(precheck_group)

        progress_group = right_bar.add_section("📊 发送进度", collapsed=True)
        progress_layout = progress_group.contentLayout()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        self.progress_label = QLabel("等待发送...")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        progress_layout.addWidget(self.progress_label)
        
        right_layout.addWidget(progress_group)
        
        log_group = right_bar.add_section("📝 日志", collapsed=True)
        log_layout = log_group.contentLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_text)
        
        right_layout.addWidget(log_group)
        
        control_group = right_bar.add_section("🎛 发送控制", collapsed=True)
        control_layout = QHBoxLayout()
        control_group.contentLayout().addLayout(control_layout)
        
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

        # 三栏引用：某栏所有面板都收起时，该栏平滑收缩为窄"芯片导轨"，
        # 腾出的宽度自动分给仍有面板展开的栏，避免大块空背板
        self.data_splitter = splitter
        self._rail_columns = [
            (left_panel, left_bar),
            (middle_panel, middle_bar),
            (right_panel, right_bar),
        ]
        self._rail_anim: Optional[QVariantAnimation] = None
        for _panel, bar in self._rail_columns:
            for sec in bar.sections:
                sec.collapsedChanged.connect(self._schedule_rail_relayout)

        # 注册所有可折叠面板，供全局精简开关遍历
        self._collapsible_groups = [
            config_group, url_group, filter_group,
            persons_group, preview_group,
            send_group, progress_group, log_group, control_group,
            recipient_group, precheck_group,
        ]
        # 保存各面板引用（自动折叠/发送时自动切换视图用）
        self.config_group = config_group
        self.url_group = url_group
        self.filter_group = filter_group
        self.persons_group = persons_group
        self.preview_group = preview_group
        self.send_group = send_group
        self.progress_group = progress_group
        self.log_group = log_group
        self.control_group = control_group
        # 加载配置后自动展开：人员列表/数据预览/发送进度/发送控制
        self._load_expand_groups = [
            persons_group, preview_group, progress_group, control_group,
        ]
        # 启动时同步一次预检缓存状态（TTL 下拉与磁盘持久化保持一致）
        self._refresh_precheck_cache_info()

    def _collapse_groups(self, *groups):
        """批量折叠指定面板（忽略 None）。"""
        for g in groups:
            if g is not None:
                g.set_collapsed(True)

    def _enter_sending_view(self):
        """点发送后自动切换视图：只保留发送控制+发送进度，其他面板果冻收起。"""
        keep = {self.control_group, self.progress_group}
        for g in self._collapsible_groups:
            should_open = g in keep
            if should_open and g.is_collapsed():
                g.set_collapsed(False, animate=True)
            elif not should_open and not g.is_collapsed():
                g.set_collapsed(True, animate=True)

    # --------------------- 空栏芯片导轨（自动收缩/恢复） ---------------------
    _RAIL_WIDTH = 126

    def _schedule_rail_relayout(self, *_args):
        """面板折叠/展开后，延迟到布局事件处理完再重算栏宽（动画起始态已确定）。"""
        QTimer.singleShot(0, self._relayout_rails)

    def _relayout_rails(self):
        """全收起的栏平滑收缩为窄芯片导轨，腾出的宽度按比例分给展开的栏。"""
        sp = self.data_splitter
        if sp is None:
            return
        start = sp.sizes()
        if len(start) != 3 or sum(start) <= 0:
            return
        total = sum(start)
        closed = [
            all(s.is_collapsed() for s in bar.sections)
            for _p, bar in self._rail_columns
        ]
        # 每栏导轨宽度按该栏最宽芯片动态计算（FlowLayout 会自动换行）
        rails = [
            max(112, bar.sizeHint().width() + 12)
            for _p, bar in self._rail_columns
        ]
        if not hasattr(self, "_rail_weights"):
            self._rail_weights = [285.0, 280.0, 285.0]
        n_closed = sum(closed)
        avail = total - sum(rails[i] for i, c in enumerate(closed) if c)
        open_weight = sum(
            self._rail_weights[i] for i, c in enumerate(closed) if not c
        )
        target = []
        for i, is_closed in enumerate(closed):
            if is_closed or open_weight <= 0:
                target.append(rails[i])
            else:
                target.append(max(rails[i], int(avail * self._rail_weights[i] / open_weight)))
        # 修正取整误差到最后一个展开栏
        diff = total - sum(target)
        if diff != 0:
            for i in range(2, -1, -1):
                if not closed[i]:
                    target[i] += diff
                    break
        if target == start:
            return
        if self._rail_anim is not None:
            try:
                self._rail_anim.stop()
            except Exception:
                pass
        anim = QVariantAnimation(self)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.setDuration(300)
        anim.setEasingCurve(QEasingCurve.Type.OutQuart)

        def on_change(t, st=list(start), tg=target):
            sp.setSizes([
                int(s + (e - s) * t) for s, e in zip(st, tg)
            ])

        def on_finish():
            # 记住当前展开栏宽度比例，供下次恢复时按原比例分配
            cur = sp.sizes()
            for i, is_closed in enumerate(closed):
                if not is_closed and cur[i] > rails[i]:
                    self._rail_weights[i] = float(cur[i])

        anim.valueChanged.connect(on_change)
        anim.finished.connect(on_finish)
        self._rail_anim = anim
        anim.start()

    def connect_signals(self):
        self.save_config_btn.clicked.connect(self.save_config)
        self.save_as_config_btn.clicked.connect(self.save_config_as)
        self.load_config_btn.clicked.connect(self.load_config)
        self.edit_config_btn.clicked.connect(self.edit_config)
        self.recent_config_list.itemDoubleClicked.connect(
            self.load_recent_config
        )
        self.recent_config_list.customContextMenuRequested.connect(
            self._on_recent_config_context_menu
        )
        self.excel_btn.clicked.connect(self.select_excel_file)
        self.reload_btn.clicked.connect(self.reload_excel_file)
        self.open_excel_btn.clicked.connect(self.open_current_excel)
        self.sheet_combo.currentIndexChanged.connect(self.on_sheet_changed)
        self.header_row_spin.valueChanged.connect(self.on_header_row_changed)
        self.local_source_radio.toggled.connect(self._on_data_source_changed)
        self.wps_settings_btn.clicked.connect(self.open_wps_settings)
        self.cloud_browse_btn.clicked.connect(self.browse_cloud_files)
        self.cloud_read_btn.clicked.connect(self.read_cloud_file)
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
        self.order_up_btn.clicked.connect(lambda: self._move_order_item(-1))
        self.order_down_btn.clicked.connect(lambda: self._move_order_item(1))
        self.send_order_list.itemChanged.connect(self._on_send_order_item_changed)

        # 联系人映射面板按钮
        self.rm_add_row_btn.clicked.connect(self._on_rm_add_row)
        self.rm_del_row_btn.clicked.connect(self._on_rm_del_row)
        self.rm_clear_btn.clicked.connect(self._on_rm_clear)
        self.rm_import_btn.clicked.connect(self._on_rm_import_table)
        self.rm_open_file_btn.clicked.connect(self._on_rm_open_mapping_file)
        self.rm_link_file_btn.clicked.connect(self._on_rm_link_mapping_file)

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

        # 序列化当前 UI 中的筛选条件（即使未点"应用筛选"也保存）
        filter_conditions = []
        for cond in self.build_filter_conditions():
            filter_conditions.append(
                {"column_name": cond.column_name, "operator": cond.operator, "value": cond.value}
            )

        excel_block = {
            "path": (
                os.path.abspath(self.current_excel_path)
                if self.data_source == "local"
                else f"wps-cloud://{self.cloud_file_id}"
            ),
            "sheet": self.sheet_combo.currentText().strip(),
            "header_row": int(self.header_row_spin.value()),
            "name_column": name_column,
            "extract_columns": list(self.selected_columns),
            "wechat_column": wechat_column,
            "source": self.data_source,
            "cloud_file_id": self.cloud_file_id if self.data_source == "wps_cloud" else "",
            "cloud_file_name": (
                self.cloud_file_name if self.data_source == "wps_cloud" else ""
            ),
        }

        return {
            "version": ConfigManager.PROFILE_VERSION,
            "excel": excel_block,
            "send": {
                "manual_recipient": self.wechat_edit.text().strip(),
                "mode": "custom",
                "send_order": self._get_send_order(),
                "interval": self.send_interval_spin.value(),
                "chat_delay": self.chat_delay_spin.value(),
                "custom_message_enabled": (
                    self.custom_msg_checkbox.isChecked()
                ),
                "custom_message": self.custom_msg_edit.toPlainText(),
                "attachment": self.attachment_edit.text().strip(),
                "auto_send": self.auto_send_check.isChecked(),
            },
            "filter_conditions": filter_conditions,
            "recipient_mapping": self._collect_recipient_mapping_from_ui(),
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
            if self.data_source == "wps_cloud":
                excel_name = self.cloud_file_name or self.cloud_file_id or "云文档"
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
        # 保存成功后自动折叠配置管理区
        self._collapse_groups(self.config_group)
        return True

    def load_config(self):
        if not self._can_load_config():
            return
        config_path = self._pick_config_file("加载配置")
        if config_path:
            self._load_config_file(config_path)

    def edit_config(self):
        """选择一个配置以"编辑模式"加载：只还原到界面，不触发 auto_send。"""
        if not self._can_load_config():
            return
        config_path = self._pick_config_file("编辑配置（不自动发送）")
        if config_path:
            self._load_config_file(config_path, trigger_auto_send=False)

    def _pick_config_file(self, dialog_title):
        start_path = (
            os.path.dirname(self.current_config_path)
            if self.current_config_path
            else str(self.config_manager.profile_dir)
        )
        config_path, _ = QFileDialog.getOpenFileName(
            self,
            dialog_title,
            start_path,
            "表格发送配置 (*.json)",
        )
        return config_path

    def load_recent_config(self, item):
        config_path = item.data(Qt.ItemDataRole.UserRole)
        if not config_path:
            return
        self._load_recent_path(config_path, trigger_auto_send=True)

    def edit_recent_config(self, item):
        config_path = item.data(Qt.ItemDataRole.UserRole)
        if not config_path:
            return
        self._load_recent_path(config_path, trigger_auto_send=False)

    def _load_recent_path(self, config_path, *, trigger_auto_send):
        if not os.path.isfile(config_path):
            self.config_manager.remove_recent(config_path)
            self.refresh_recent_configs()
            QMessageBox.warning(self, "配置不存在", "该配置文件已被移动或删除")
            return
        self._load_config_file(
            config_path, trigger_auto_send=trigger_auto_send)

    def _on_recent_config_context_menu(self, pos):
        item = self.recent_config_list.itemAt(pos)
        if item is None:
            return
        menu = QMenu(self)
        act_load = menu.addAction("加载配置（按配置自动发送）")
        act_edit = menu.addAction("编辑配置（不自动发送）")
        chosen = menu.exec(
            self.recent_config_list.viewport().mapToGlobal(pos))
        if chosen == act_load:
            self.load_recent_config(item)
        elif chosen == act_edit:
            self.edit_recent_config(item)

    def _can_load_config(self):
        if self.excel_worker and self.excel_worker.isRunning():
            QMessageBox.warning(self, "请稍候", "Excel 正在读取，暂时不能加载配置")
            return False
        if self.worker and self.worker.isRunning():
            QMessageBox.warning(self, "请稍候", "正在发送消息，暂时不能加载配置")
            return False
        return True

    def _load_config_file(self, config_path, *, trigger_auto_send=True):
        if not self._can_load_config():
            return False

        try:
            profile = self.config_manager.load_profile(config_path)
        except ConfigError as error:
            QMessageBox.warning(self, "加载失败", str(error))
            self.log(f"✗ 配置加载失败: {error}")
            return False

        excel_settings = profile["excel"]
        source = excel_settings.get("source", "local")

        self._set_current_config(config_path)
        self.config_manager.add_recent(self.current_config_path)
        self.refresh_recent_configs()
        if trigger_auto_send:
            self.log(
                f"正在加载配置: {os.path.basename(self.current_config_path)}"
            )
        else:
            self.log(
                f"正在以编辑模式加载配置（不会自动发送）: "
                f"{os.path.basename(self.current_config_path)}"
            )

        if source == "wps_cloud":
            file_token = excel_settings.get("cloud_file_id", "").strip()
            if not file_token:
                QMessageBox.warning(self, "加载失败", "该配置缺少云文档 ID")
                return False
            file_name = excel_settings.get("cloud_file_name", "")
            self._set_data_source("wps_cloud", refresh=False)
            self.cloud_file_edit.setText(file_token)
            self.cloud_file_id = file_token
            self.cloud_file_name = file_name
            self.pending_config = {
                "profile": profile,
                "excel_path": f"wps-cloud://{file_token}",
                "suppress_auto_send": not trigger_auto_send,
            }
            sheet_name = excel_settings.get("sheet") or None
            if not self._load_cloud_data(file_token, file_name, sheet_name):
                self.pending_config = None
                return False
            return True

        excel_path = excel_settings["path"]
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

        self._set_data_source("local", refresh=False)
        self.pending_config = {
            "profile": profile,
            "excel_path": os.path.abspath(excel_path),
            "suppress_auto_send": not trigger_auto_send,
        }
        sheet_name = excel_settings.get("sheet") or None
        if not self._load_excel_data(excel_path, sheet_name):
            self.pending_config = None
            return False

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
        # 编辑模式加载时抑制本次 auto_send（标志随 pending_config 同生命周期）
        suppress_auto_send = bool(
            self.pending_config.get("suppress_auto_send", False))
        self.pending_config = None
        excel_settings = profile["excel"]
        send_settings = profile["send"]
        missing_settings = []

        saved_sheet = excel_settings.get("sheet", "")
        if saved_sheet and saved_sheet != self.current_sheet:
            missing_settings.append(
                f"Sheet“{saved_sheet}”不存在，已使用“{self.current_sheet}”"
            )

        # 表头行在 on_excel_read_result 已按配置静默应用；此处仅做越界提示
        try:
            saved_header_row = int(excel_settings.get("header_row", 1))
        except (TypeError, ValueError):
            saved_header_row = 1
        total_rows = len(self.table_data or [])
        if saved_header_row < 1 or saved_header_row > total_rows:
            missing_settings.append(
                f"表头行 {saved_header_row} 超出范围（共 {total_rows} 行），"
                f"已回退到第 {self.header_row_spin.value()} 行"
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
        self.send_interval_spin.setValue(send_settings.get("interval", 0.5))
        self.chat_delay_spin.setValue(
            send_settings.get("chat_delay", 0.3)
        )
        self.custom_msg_edit.setPlainText(
            send_settings.get("custom_message", "")
        )
        self.attachment_edit.setText(
            send_settings.get("attachment", "") or ""
        )
        self.auto_send_check.setChecked(
            bool(send_settings.get("auto_send", False))
        )

        # 联系人映射：旧配置无此字段时按禁用+空表初始化（向后兼容）
        rm_cfg = profile.get("recipient_mapping") or {}
        self.recipient_mapping_enabled_check.setChecked(
            bool(rm_cfg.get("enabled", False))
        )
        self.recipient_mapping_default_edit.setText(
            str(rm_cfg.get("default_recipient", "") or "")
        )
        rm_file = str(rm_cfg.get("mapping_file", "") or "").strip()
        self._rm_linked_file = rm_file
        if rm_file:
            self.rm_link_file_btn.setToolTip(f"已关联: {rm_file}")
        else:
            self.rm_link_file_btn.setToolTip("关联一个映射表文件路径")
        # 把 mappings 列表（每项 source_value -> recipients）填到表格
        merged = {}
        for m in (rm_cfg.get("mappings") or []):
            src = str(m.get("source_value", "") or "").strip()
            recips_raw = m.get("recipients") or []
            if not isinstance(recips_raw, list):
                continue
            recips = [str(r).strip() for r in recips_raw if str(r).strip()]
            if src and recips:
                merged[src] = recips
        self._populate_recipient_mapping_table(merged)

        # 发送顺序：优先读新字段 send_order；旧配置从 mode 迁移
        custom_enabled = bool(send_settings.get("custom_message_enabled", False))
        attachment_path = self.attachment_edit.text().strip()
        raw_order = send_settings.get("send_order")
        valid_keys = {k for k, _ in self._ORDER_ITEMS}
        if isinstance(raw_order, list) and raw_order:
            order = [k for k in raw_order if k in valid_keys]
            if custom_enabled and "custom" not in order:
                order.append("custom")
            if attachment_path and "attachment" not in order:
                order.append("attachment")
        else:
            mode = send_settings.get("mode", "text")
            if mode == "image":
                order = ["image"]
            elif mode == "image_text":
                order = ["image", "text"]
            else:
                order = ["text"]
            if custom_enabled:
                order.append("custom")
            if attachment_path:
                order.append("attachment")
        # 勾选了但内容为空的不发送（防止误发空附件步骤）
        if not custom_enabled and "custom" in order:
            order.remove("custom")
        if not attachment_path and "attachment" in order:
            order.remove("attachment")
        self._set_send_order(order)
        self.custom_msg_checkbox.setChecked("custom" in order)

        # 恢复筛选条件：先清空旧 UI 行，再按配置逐条重建
        saved_filters = profile.get("filter_conditions", [])
        # 直接清空旧筛选行（不用 clear_filter_conditions 避免其 refresh 副作用）
        while self.filter_conditions_layout.count() > 0:
            item = self.filter_conditions_layout.takeAt(0)
            if item and item.widget():
                item.widget().deleteLater()
        self.filter_conditions = []

        str_headers = [str(h) for h in (self.headers or [])]
        for cond in saved_filters:
            col = cond.get("column_name", "")
            op = cond.get("operator", "")
            val = cond.get("value", "")
            if not (col and op):
                continue
            # 内联创建筛选行（不调用 add_filter_condition，绕过 processor 守卫）
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(2, 2, 2, 2)

            col_combo = QComboBox()
            col_combo.addItems(str_headers)
            if col in str_headers:
                col_combo.setCurrentText(col)
            row_layout.addWidget(col_combo)

            op_combo = QComboBox()
            op_combo.addItems([name for name, _ in self.get_available_operators()])
            for i, (_, key) in enumerate(self.get_available_operators()):
                if key == op:
                    op_combo.setCurrentIndex(i)
                    break
            row_layout.addWidget(op_combo)

            value_edit = QLineEdit(val)
            value_edit.setPlaceholderText("条件值")
            row_layout.addWidget(value_edit)

            remove_btn = QPushButton("×")
            remove_btn.setStyleSheet("background-color: #f44336; color: white; padding: 2px 6px;")
            row_layout.addWidget(remove_btn)
            remove_btn.clicked.connect(
                lambda checked=False, w=row_widget: self.remove_filter_condition(w)
            )

            col_combo.currentTextChanged.connect(self.on_filter_change)
            op_combo.currentTextChanged.connect(self.on_filter_change)
            value_edit.textChanged.connect(self.on_filter_change)

            self.filter_conditions_layout.addWidget(row_widget)

        self.apply_filter_btn.setEnabled(self.filter_conditions_layout.count() > 0)

        can_load_data = (
            name_column in self.headers
            and bool(self.selected_columns)
        )
        if can_load_data:
            self.load_data()
            self.log("配置已应用，并已自动加载人员数据")
        else:
            self.log("⚠ 配置已部分应用，请重新选择缺失的列")

        # load_data 创建 processor 后再应用筛选，否则 apply_filter 因 processor=None 直接返回
        if saved_filters:
            self.apply_filter()

        if missing_settings:
            QMessageBox.warning(
                self,
                "配置部分失效",
                "\n".join(missing_settings),
            )

        # auto_send：配置要求加载后自动全选筛选人员并直接发送。
        # 延迟到本轮 UI 处理完成后再执行，确保列表已渲染、弹窗已关闭。
        # 编辑模式加载（suppress_auto_send）只把配置还原到界面供修改，不触发发送。
        if can_load_data and bool(send_settings.get("auto_send", False)) \
                and suppress_auto_send:
            self.log(
                "⚙ 该配置勾选了「加载后自动发送」，当前为编辑模式，已跳过"
                "自动发送；修改后请点「保存配置」"
            )

        if can_load_data and bool(send_settings.get("auto_send", False)) \
                and not suppress_auto_send:
            def _auto_start_send():
                try:
                    if not self.processor:
                        return
                    self.persons_list.selectAll()
                    selected = self.persons_list.selectedItems()
                    if selected:
                        self.send_data(auto=True)
                    else:
                        self.log("⚡ 自动发送：筛选结果为空，未发送")
                except Exception as exc:
                    self.log(f"⚡ 自动发送启动失败: {exc}")

            QTimer.singleShot(300, _auto_start_send)
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
        col_combo.addItems([str(h) for h in headers])
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
        checked = state == Qt.CheckState.Checked.value
        self.custom_msg_edit.setEnabled(checked)
        # 同步发送顺序列表里的“自定义消息”勾选状态
        self._set_order_item_checked("custom", checked)

    def _on_pick_attachment(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择附加文件(文件/图片)",
            "",
            "常用文件 (*.png *.jpg *.jpeg *.gif *.bmp *.webp *.pdf *.doc *.docx *.xls *.xlsx *.ppt *.pptx *.txt *.zip *.rar *.7z);;所有文件 (*.*)",
        )
        if path:
            self.attachment_edit.setText(path)
            # 选中文件自动勾选发送顺序里的“附加文件”
            self._set_order_item_checked("attachment", True)

    def _on_clear_attachment(self):
        self.attachment_edit.clear()
        self._set_order_item_checked("attachment", False)

    # ------------------------- 联系人映射面板 -------------------------
    _RM_SEPARATORS = ("/", ";", "、", "，", ",")

    @classmethod
    def _split_rm_values(cls, raw):
        """按 / ; 、 ， , 拆分多值，去空白和空串。

        引号包裹的整串视为一个完整值：其内部的分隔符（含逗号）不生效，
        并去掉两端引号——如 `"孙中枢,孙中枢"` 应保留为单个 `孙中枢,孙中枢`，
        而不是被拆成 `孙中枢` 与 `孙中枢"`（与全量默认映射导入
        _plain_split_values 行为一致，名字本身可能含逗号）。
        """
        if not raw:
            return []
        text = str(raw).strip()
        if not text:
            return []
        parts = []
        buf = []
        in_quote = False
        for ch in text:
            if ch == '"':
                in_quote = not in_quote
                buf.append(ch)
            elif ch in cls._RM_SEPARATORS and not in_quote:
                parts.append("".join(buf).strip())
                buf = []
            else:
                buf.append(ch)
        parts.append("".join(buf).strip())
        out = []
        for p in parts:
            p = p.strip()
            if p.startswith('"') and p.endswith('"') and len(p) >= 2:
                p = p[1:-1].strip()
            if p:
                out.append(p)
        return out

    def _on_rm_add_row(self):
        """添加一行空白映射，方便手动填写。"""
        table = self.recipient_mapping_table
        row = table.rowCount()
        table.insertRow(row)
        table.setItem(row, 0, QTableWidgetItem(""))
        table.setItem(row, 1, QTableWidgetItem(""))
        table.editItem(table.item(row, 0))

    def _on_rm_del_row(self):
        """删除选中的所有行。"""
        table = self.recipient_mapping_table
        rows = sorted(
            {idx.row() for idx in table.selectedIndexes()},
            reverse=True,
        )
        for r in rows:
            table.removeRow(r)

    def _on_rm_clear(self):
        """清空映射表所有行。"""
        self.recipient_mapping_table.setRowCount(0)

    def _on_rm_link_mapping_file(self):
        """关联一个映射表文件（.xlsx/.csv），保存到当前 profile 的 mapping_file 字段。"""
        start_dir = ""
        cur = self._rm_linked_file
        if cur and os.path.isdir(os.path.dirname(cur)):
            start_dir = os.path.dirname(cur)
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择映射表文件",
            start_dir,
            "映射表 (*.xlsx *.xls *.csv);;所有文件 (*.*)",
        )
        if not path:
            return
        self._rm_linked_file = os.path.abspath(path)
        self.rm_link_file_btn.setToolTip(f"已关联: {self._rm_linked_file}")
        self.log(f"已关联映射表文件: {self._rm_linked_file}")

    @property
    def _rm_linked_file(self):
        """当前 profile 关联的映射表文件路径（内存态，保存配置时持久化）。"""
        return getattr(self, "_rm_linked_file_path", "") or ""

    @_rm_linked_file.setter
    def _rm_linked_file(self, path):
        self._rm_linked_file_path = path or ""

    def _on_rm_open_mapping_file(self):
        """用系统默认编辑器（Excel/WPS）打开已关联的映射表文件。

        若未关联则先弹出文件选择框，记下路径再打开。
        """
        path = self._rm_linked_file
        if not path or not os.path.isfile(path):
            # 未关联或文件丢失 → 让用户先选一个
            start_dir = ""
            if path and os.path.isdir(os.path.dirname(path)):
                start_dir = os.path.dirname(path)
            picked, _ = QFileDialog.getOpenFileName(
                self,
                "选择要打开的映射表文件",
                start_dir,
                "映射表 (*.xlsx *.xls *.csv);;所有文件 (*.*)",
            )
            if not picked:
                return
            path = os.path.abspath(picked)
            self._rm_linked_file = path
            self.rm_link_file_btn.setToolTip(f"已关联: {path}")
        try:
            # Windows: os.startfile；其他平台兜底用 QDesktopServices
            if hasattr(os, "startfile"):
                os.startfile(path)
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(path))
            self.log(f"已用系统默认程序打开映射表: {path}")
        except OSError as exc:
            QMessageBox.warning(
                self, "打开失败", f"无法打开映射表文件:\n{path}\n\n错误: {exc}"
            )

    def _on_rm_import_table(self):
        """从 .xlsx/.csv 导入映射关系，合并到当前表格。

        文件两列结构：
          第一列 = 筛选值（可含 / 或 ; 分隔多个值，多个值共享同一组接收人）
          第二列 = 接收人（可含 / 或 ; 分隔多人）

        合并规则：同 source_value 已存在则覆盖，不存在则新增。
        """
        start_dir = ""
        cur = self._rm_linked_file
        if cur and os.path.isdir(os.path.dirname(cur)):
            start_dir = os.path.dirname(cur)
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择映射表文件",
            start_dir,
            "映射表 (*.xlsx *.xls *.csv);;所有文件 (*.*)",
        )
        if not path:
            return

        # 复用项目内的 read_one_sheet 读取（不依赖 pandas 直接调用 Qt 线程）
        try:
            from modules.wps_cloud import read_one_sheet, read_sheet_names
        except ImportError as exc:
            QMessageBox.critical(self, "导入失败", f"缺少必要模块: {exc}")
            return

        # read_one_sheet 对 sheet_name=None 会返回 None（None 不在 sheet 列表里），
        # 必须先取 sheet 名称列表，再显式指定第一个 sheet 读取
        try:
            names = read_sheet_names(path, log_fn=self.log)
            if not names:
                QMessageBox.warning(
                    self, "导入失败",
                    "无法读取工作表名称，文件可能为空或格式不支持"
                )
                return
            rows = read_one_sheet(path, names[0], log_fn=self.log)
        except Exception as exc:
            QMessageBox.warning(
                self, "导入失败", f"读取映射表失败:\n{path}\n\n错误: {exc}"
            )
            return

        if not rows:
            QMessageBox.warning(self, "导入失败", "映射表为空或无有效数据")
            return

        # 解析为 dict[source_value] -> list[recipients]，便于合并覆盖
        merged = {}
        # 先把当前 UI 中的映射读出来
        for r in range(self.recipient_mapping_table.rowCount()):
            src_item = self.recipient_mapping_table.item(r, 0)
            rec_item = self.recipient_mapping_table.item(r, 1)
            if not src_item:
                continue
            for s in self._split_rm_values(src_item.text()):
                recips = self._split_rm_values(rec_item.text() if rec_item else "")
                if s and recips:
                    merged[s] = recips

        # 合并导入的数据
        imported_count = 0
        for row in rows:
            if not row or len(row) < 2:
                continue
            srcs = self._split_rm_values(row[0])
            recips = self._split_rm_values(row[1])
            if not srcs or not recips:
                continue
            for s in srcs:
                merged[s] = list(recips)  # 覆盖式合并
                imported_count += 1

        # 写回表格
        self._populate_recipient_mapping_table(merged)

        # 记下文件路径，便于后续"打开"
        self._rm_linked_file = os.path.abspath(path)
        self.rm_link_file_btn.setToolTip(f"已关联: {self._rm_linked_file}")

        self.log(
            f"导入完成：合并后共 {len(merged)} 条映射"
            f"（本次新增/覆盖 {imported_count} 条），来源: {path}"
        )

    def _populate_recipient_mapping_table(self, merged_dict):
        """把 dict[source_value -> list[recipients]] 写到表格，每行一个 source_value。"""
        table = self.recipient_mapping_table
        table.setRowCount(0)
        for src, recips in merged_dict.items():
            row = table.rowCount()
            table.insertRow(row)
            table.setItem(row, 0, QTableWidgetItem(src))
            table.setItem(row, 1, QTableWidgetItem(" / ".join(recips)))

    def _collect_recipient_mapping_from_ui(self):
        """从 UI 收集联系人映射配置，返回与 profile JSON 一致的 dict。

        返回 dict 中额外有 value_to_recipients 字段，方便发送逻辑直接使用。
        """
        enabled = self.recipient_mapping_enabled_check.isChecked()
        default_recipient = self.recipient_mapping_default_edit.text().strip()

        # UI 表格每行允许 raw 多值（含分隔符）；保存时按分隔符拆分两列，
        # 每行的多个 source_value 共享同一组 recipients，最终序列化为扁平的 mappings。
        merged = {}
        for r in range(self.recipient_mapping_table.rowCount()):
            src_item = self.recipient_mapping_table.item(r, 0)
            rec_item = self.recipient_mapping_table.item(r, 1)
            if not src_item:
                continue
            srcs = self._split_rm_values(src_item.text())
            recips = self._split_rm_values(rec_item.text() if rec_item else "")
            if not srcs or not recips:
                continue
            for s in srcs:
                # 同 source_value 出现多次时，后者覆盖前者
                merged[s] = list(recips)

        mappings = [
            {"source_value": s, "recipients": rs}
            for s, rs in merged.items()
        ]

        return {
            "enabled": enabled,
            "default_recipient": default_recipient,
            "mapping_file": self._rm_linked_file,
            "mappings": mappings,
            # 发送逻辑辅助字段（不持久化，仅内存使用）
            "value_to_recipients": merged if enabled else {},
        }

    def _resolve_recipients_for_person(
        self, name, value_to_recipients, default_recipient, mapping_enabled,
        fallback_recipient,
    ):
        """统一的接收人解析：映射优先 → 兜底 → fallback。

        fallback_recipient 通常来自原"筛选列"逻辑（wechat_mapping / manual_recipient / name）。
        返回 list[str]（一对多时多个）。
        """
        if mapping_enabled and name in value_to_recipients:
            return list(value_to_recipients[name])
        if mapping_enabled and default_recipient:
            return [default_recipient]
        return [fallback_recipient] if fallback_recipient else []

    def _profile_mapping_to_dict(self, profile):
        """把 profile 里的 recipient_mapping 展开为 (enabled, value_to_recipients, default_recipient, mapping_file)。"""
        rm = profile.get("recipient_mapping") or {}
        enabled = bool(rm.get("enabled", False))
        default_recipient = str(rm.get("default_recipient", "") or "").strip()
        mapping_file = str(rm.get("mapping_file", "") or "").strip()
        value_to_recipients = {}
        if enabled:
            for m in (rm.get("mappings") or []):
                src = str(m.get("source_value", "") or "").strip()
                recips_raw = m.get("recipients") or []
                if not isinstance(recips_raw, list):
                    continue
                recips = [str(r).strip() for r in recips_raw if str(r).strip()]
                if src and recips:
                    value_to_recipients[src] = recips
        return enabled, value_to_recipients, default_recipient, mapping_file

    def _resolve_mapping_for_send(self):
        """统一解析本次发送应使用的联系人映射。

        优先用 UI 联系人映射面板当前状态；若面板未配置任何映射
        （手动选 Excel 发送、未先加载配置的场景），则回退从"当前配置
        /最近使用的配置"恢复已保存的联系人映射，确保手动选人/全选同
        自动发送一样走映射关系。
        返回 (enabled, value_to_recipients, default_recipient, mapping_file, from_ui)。
        """
        rm_cfg = self._collect_recipient_mapping_from_ui()
        enabled = rm_cfg.get("enabled", False)
        value_to_recipients = rm_cfg.get("value_to_recipients", {})
        default_recipient = rm_cfg.get("default_recipient", "")
        mapping_file = str(rm_cfg.get("mapping_file", "") or "")
        from_ui = True

        if enabled and value_to_recipients:
            return enabled, value_to_recipients, default_recipient, mapping_file, from_ui

        # 面板无有效映射 → 尝试从配置恢复。
        # 显式加载了配置（current_config_path 非空）时**只**认当前配置自身的映射：
        # 若用户关闭了该配置的「启用映射表」，应继续走到全量默认映射兜底，而不是
        # 被「最近使用的配置」里其他配置的映射抢占（否则全量兜底永远轮不到）。
        # 仅当未加载任何配置（纯手动选 Excel 发送）时才从最近使用的配置恢复映射。
        candidates = []
        if self.current_config_path:
            candidates.append(self.current_config_path)
        else:
            for p in self.config_manager.get_recent_profiles():
                candidates.append(p)

        for path in candidates:
            try:
                profile = self.config_manager.load_profile(path)
            except Exception:
                continue
            rm = profile.get("recipient_mapping") or {}
            if not rm.get("enabled"):
                continue
            pe, p2r, pdefault, pfile = self._profile_mapping_to_dict(profile)
            if pe and p2r:
                self.log(
                    f"⚙ 手动发送：UI 未配置联系人映射，已从配置恢复 "
                    f"{len(p2r)} 条映射（{os.path.basename(path)}）"
                )
                return pe, p2r, pdefault, pfile, False

        # 最后兜底：全量默认映射（仅当 UI 面板、当前/最近配置都未启用映射时）
        try:
            gs_settings = self.config_manager.load_global_settings()
        except Exception:
            gs_settings = {}
        g_dm = gs_settings.get("default_mapping") or {}
        if g_dm.get("enabled"):
            g_v2r = {}
            for m in (g_dm.get("mappings") or []):
                src = str(m.get("source_value", "") or "").strip()
                recips = [str(r).strip() for r in (m.get("recipients") or []) if str(r).strip()]
                if src and recips:
                    g_v2r[src] = recips
            g_default = str(g_dm.get("default_recipient", "") or "").strip()
            if g_v2r or g_default:
                self.log(
                    f"⚙ 未匹配到配置映射，已应用「全量默认映射」"
                    f"{len(g_v2r)} 条作为最后兜底"
                )
                return True, g_v2r, g_default, "", False
        return enabled, value_to_recipients, default_recipient, mapping_file, from_ui

    # ------------------------- 发送预检缓存 -------------------------
    def _on_precheck_ttl_changed(self, _idx=None):
        """缓存有效期下拉变化：写入 WeChatSender 并持久化。"""
        try:
            from modules.wechat_sender import WeChatSender
            sender = WeChatSender.shared_instance()
            ttl = int(self.precheck_ttl_combo.currentData() or 0)
            sender.set_precheck_ttl(ttl)
            self.log(
                f"发送预检缓存有效期已设置为: {self.precheck_ttl_combo.currentText()}"
            )
            self._refresh_precheck_cache_info()
        except Exception as exc:
            self.log(f"⚠ 设置预检缓存有效期失败: {exc}")

    def _on_clear_precheck_cache(self):
        """手动清空发送预检缓存。"""
        try:
            from modules.wechat_sender import WeChatSender
            sender = WeChatSender.shared_instance()
            n = sender.clear_precheck_cache()
            self.log(f"🧹 已清除发送预检缓存 {n} 条")
            self._refresh_precheck_cache_info()
        except Exception as exc:
            self.log(f"⚠ 清除预检缓存失败: {exc}")

    def _refresh_precheck_cache_info(self):
        """刷新缓存状态标签 + 下拉框与持久化 TTL 同步（不触发变更信号）。"""
        try:
            from modules.wechat_sender import WeChatSender
            sender = WeChatSender.shared_instance()
            count, ttl = sender.get_precheck_cache_stats()
            ttl_text = "永久" if ttl <= 0 else f"{ttl} 秒"
            self.precheck_cache_info.setText(f"当前缓存: {count} 条 · 有效期: {ttl_text}")
            idx = self.precheck_ttl_combo.findData(ttl)
            if idx >= 0 and idx != self.precheck_ttl_combo.currentIndex():
                self.precheck_ttl_combo.blockSignals(True)
                self.precheck_ttl_combo.setCurrentIndex(idx)
                self.precheck_ttl_combo.blockSignals(False)
        except Exception:
            self.precheck_cache_info.setText("预检缓存: -")

    # ------------------------- 发送顺序 -------------------------
    def _get_send_order(self):
        """按列表行顺序返回已勾选的内容类型 key 列表。"""
        order = []
        for i in range(self.send_order_list.count()):
            item = self.send_order_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                order.append(item.data(Qt.ItemDataRole.UserRole))
        return order

    def _find_order_item(self, key):
        for i in range(self.send_order_list.count()):
            item = self.send_order_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == key:
                return i, item
        return -1, None

    def _set_order_item_checked(self, key, checked):
        _idx, item = self._find_order_item(key)
        if item is None:
            return
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        if item.checkState() != state:
            self.send_order_list.blockSignals(True)
            try:
                item.setCheckState(state)
            finally:
                self.send_order_list.blockSignals(False)
        if key == "custom":
            self.custom_msg_edit.setEnabled(checked)

    def _set_send_order(self, order_keys):
        """按给定顺序重建列表行并设置勾选状态。"""
        label_map = dict(self._ORDER_ITEMS)
        self.send_order_list.blockSignals(True)
        try:
            self.send_order_list.clear()
            enabled = set(order_keys)
            ordered = list(order_keys)
            # 未包含的类型追加在后面（未勾选），保证四项始终可见可勾选
            for key, _label in self._ORDER_ITEMS:
                if key not in enabled and key not in ordered:
                    ordered.append(key)
            for key in ordered:
                item = QListWidgetItem(label_map.get(key, key))
                item.setData(Qt.ItemDataRole.UserRole, key)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(
                    Qt.CheckState.Checked if key in enabled else Qt.CheckState.Unchecked
                )
                self.send_order_list.addItem(item)
        finally:
            self.send_order_list.blockSignals(False)
        self.custom_msg_edit.setEnabled("custom" in enabled)

    def _move_order_item(self, delta):
        row = self.send_order_list.currentRow()
        if row < 0:
            return
        target = row + delta
        if target < 0 or target >= self.send_order_list.count():
            return
        item = self.send_order_list.takeItem(row)
        self.send_order_list.insertItem(target, item)
        self.send_order_list.setCurrentRow(target)

    def _on_send_order_item_changed(self, item):
        # 自定义消息勾选状态联动编辑框可用性 + 顶部复选框
        if item.data(Qt.ItemDataRole.UserRole) == "custom":
            checked = item.checkState() == Qt.CheckState.Checked
            self.custom_msg_edit.setEnabled(checked)
            self.custom_msg_checkbox.blockSignals(True)
            try:
                self.custom_msg_checkbox.setChecked(checked)
            finally:
                self.custom_msg_checkbox.blockSignals(False)

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
            if self.persons_list.count() == 0:
                self.log("⚠ 筛选结果为空，没有符合条件的人员，请调整筛选条件")
                QMessageBox.information(
                    self,
                    "筛选结果为空",
                    "当前筛选条件下没有匹配的人员，请调整或清除筛选条件后重试。",
                )
        else:
            self.log("清除筛选条件，显示全部人员")

        # 应用筛选后自动展开人员列表面板（筛选后面板可能处于折叠状态）
        if self.persons_group.is_collapsed():
            self.persons_group.set_collapsed(False)

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
            # 选中人员后自动展开数据预览面板（若当前折叠）
            if self.preview_group.is_collapsed():
                self.preview_group.set_collapsed(False, animate=True)
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
        if self.data_source == "wps_cloud":
            if not self.cloud_file_id:
                QMessageBox.warning(self, "警告", "请先填写在线文档 ID")
                return
            if self.excel_worker and self.excel_worker.isRunning():
                self.log("云文档正在读取，请稍候...")
                return
            self.log("重新读取云文档...")
            self._load_cloud_data(
                self.cloud_file_id,
                self.cloud_file_name,
                self.current_sheet,
            )
            return

        if not getattr(self, 'current_excel_path', None):
            QMessageBox.warning(self, "警告", "请先选择Excel文件")
            return

        if self.excel_worker and self.excel_worker.isRunning():
            self.log("Excel文件正在读取，请稍候...")
            return

        self.log("重新读取Excel文件...")
        self._load_excel_data(self.current_excel_path)

    def open_current_excel(self):
        if self.data_source == "wps_cloud":
            if not self.cloud_file_id:
                return
            url = f"https://www.kdocs.cn/office/{self.cloud_file_id}"
            try:
                import webbrowser
                webbrowser.open(url, new=1)
                self.log(f"已在浏览器打开云文档: {url}")
            except Exception as e:
                self.log(f"✗ 打开云文档失败: {e}")
            return
        if hasattr(self, 'current_excel_path') and self.current_excel_path:
            try:
                os.startfile(self.current_excel_path)
                self.log(f"已打开文件: {self.current_excel_path}")
            except Exception as e:
                self.log(f"✗ 打开文件失败: {e}")
                QMessageBox.warning(self, "打开失败", f"无法打开文件: {e}")

    def _set_excel_loading(self, loading):
        is_cloud = self.data_source == "wps_cloud"
        self.excel_btn.setEnabled(not loading and not is_cloud)
        has_file = bool(self.current_excel_path)
        has_data = bool(self.headers and self.table_data)
        self.reload_btn.setEnabled(
            not loading and (has_file if not is_cloud else bool(self.cloud_file_id))
        )
        self.open_excel_btn.setEnabled(
            not loading and has_file and not is_cloud
        )
        self.sheet_combo.setEnabled(
            not loading and self.sheet_combo.count() > 0
        )
        self.header_row_spin.setEnabled(not loading and bool(self.table_data))
        self.name_column_combo.setEnabled(not loading and has_data)
        self.extract_columns_btn.setEnabled(not loading and has_data)
        self.wechat_column_combo.setEnabled(not loading and has_data)
        self.load_data_btn.setEnabled(not loading and has_data)
        self.save_config_btn.setEnabled(not loading and has_data)
        self.save_as_config_btn.setEnabled(not loading and has_data)
        self.load_config_btn.setEnabled(not loading)
        self.edit_config_btn.setEnabled(not loading)
        self.recent_config_list.setEnabled(not loading)
        if is_cloud:
            self.cloud_file_edit.setEnabled(not loading)
            self.cloud_browse_btn.setEnabled(not loading)
            self.cloud_read_btn.setEnabled(not loading)

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

    # ------------------------------------------------------------------
    # 金山文档在线表格
    # ------------------------------------------------------------------

    def _on_data_source_changed(self):
        source = "wps_cloud" if self.cloud_source_radio.isChecked() else "local"
        self._set_data_source(source)

    def _set_data_source(self, source, refresh=True):
        self.data_source = source
        is_cloud = source == "wps_cloud"

        self.local_source_radio.blockSignals(True)
        self.cloud_source_radio.blockSignals(True)
        try:
            self.local_source_radio.setChecked(not is_cloud)
            self.cloud_source_radio.setChecked(is_cloud)
        finally:
            self.local_source_radio.blockSignals(False)
            self.cloud_source_radio.blockSignals(False)

        self.local_row_widget.setVisible(not is_cloud)
        self.cloud_row_widget.setVisible(is_cloud)
        self.cloud_status_label.setVisible(is_cloud)
        if is_cloud:
            configured = self.wps_store.is_configured()
            self.cloud_file_edit.setEnabled(configured)
            self.cloud_browse_btn.setEnabled(configured)
            self.cloud_read_btn.setEnabled(configured)
        else:
            # 切回本地模式必须恢复按钮启用状态（云文档模式下 excel_btn 会被禁用）
            loading = bool(self.excel_worker and self.excel_worker.isRunning())
            self._set_excel_loading(loading)
        self._update_cloud_status()
        if refresh and is_cloud and not self.wps_store.is_configured():
            self.log("当前为云文档来源，请先点击“Cookie 凭证设置”粘贴浏览器里的 wps_sid")

    def _update_cloud_status(self):
        if not hasattr(self, "cloud_status_label"):
            return
        if self.wps_store.is_configured():
            self.cloud_status_label.setText(
                "登录凭证已配置（wps_sid 失效时在“Cookie 凭证设置”重新粘贴）"
            )
            self.cloud_status_label.setStyleSheet("color: #2E7D32; font-size: 11px;")
        else:
            self.cloud_status_label.setText(
                "未配置登录凭证 → 点击“Cookie 凭证设置”粘贴 wps_sid"
            )
            self.cloud_status_label.setStyleSheet("color: #C62828; font-size: 11px;")

    def open_wps_settings(self):
        dialog = WpsCredentialDialog(self.wps_store, self)
        dialog.credentials_changed.connect(self._on_credentials_changed)
        dialog.exec()

    def _on_credentials_changed(self):
        self._update_cloud_status()
        if self.data_source == "wps_cloud":
            configured = self.wps_store.is_configured()
            self.cloud_file_edit.setEnabled(configured)
            self.cloud_browse_btn.setEnabled(configured)
            self.cloud_read_btn.setEnabled(configured)

    def read_cloud_file(self):
        if self.excel_worker and self.excel_worker.isRunning():
            self.log("云文档正在读取，请稍候...")
            return
        raw = self.cloud_file_edit.text().strip()
        try:
            file_token = parse_file_input(raw)
        except WpsCloudError as exc:
            QMessageBox.warning(self, "文档 ID 无效", str(exc))
            return
        # 直接粘贴 ID 时文件名未知，清空旧文件名
        if raw == file_token:
            self.cloud_file_name = ""
        self.pending_config = None
        self._load_cloud_data(file_token, self.cloud_file_name)

    def _load_cloud_data(self, file_token, file_name="", sheet_name=None):
        if self.excel_worker and self.excel_worker.isRunning():
            self.log("文件正在读取，已忽略重复请求")
            return False
        if not self.wps_store.is_configured():
            QMessageBox.warning(
                self, "未配置凭证",
                "请先点击“Cookie 凭证设置”粘贴浏览器里的 wps_sid",
            )
            return False

        self.cloud_file_id = file_token
        self.log(f"开始读取金山云文档: {file_name or file_token}")
        self._set_excel_loading(True)

        worker = CloudExcelReadWorker(file_token, file_name, sheet_name)
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

    def browse_cloud_files(self):
        if not self.wps_store.is_configured():
            QMessageBox.warning(
                self, "未配置凭证",
                "请先点击“Cookie 凭证设置”粘贴浏览器里的 wps_sid",
            )
            return
        if self.excel_worker and self.excel_worker.isRunning():
            self.log("请等待当前读取完成")
            return
        self.log("正在获取云文档列表…")
        self.cloud_browse_btn.setEnabled(False)
        worker = WpsFileListWorker(self.wps_store)
        self._cloud_list_worker = worker
        worker.signals.result.connect(self.on_cloud_list_result)
        worker.signals.error.connect(self.on_cloud_list_error)
        worker.signals.log.connect(self.log)
        worker.finished.connect(worker.deleteLater)
        worker.start()

    def on_cloud_list_result(self, result):
        self.cloud_browse_btn.setEnabled(True)
        files = result.get("files") or []
        if not files:
            QMessageBox.information(
                self, "云文档列表",
                "未获取到在线表格文件（账号下需要存在 xlsx/xls/et 在线表格）",
            )
            return
        dialog = CloudFilePickerDialog(files, self)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.selected:
            name, file_id, group_id = dialog.selected
            token = f"{group_id}:{file_id}" if group_id else file_id
            self.cloud_file_edit.setText(token)
            self.cloud_file_id = token
            self.cloud_file_name = name
            self.pending_config = None
            self._load_cloud_data(token, name)

    def on_cloud_list_error(self, error):
        self.cloud_browse_btn.setEnabled(True)
        self.log(f"✗ 获取云文档列表失败: {error}")
        QMessageBox.warning(self, "获取列表失败", error)

    def on_excel_read_result(self, worker, result):
        if worker is not self.excel_worker:
            return

        self.sheet_names = result['sheet_names']
        self.current_sheet = result['current_sheet']
        self.table_data = result['data']
        # 表头行：加载配置时用配置中的行号；手动切换 Sheet/重新读取时保留用户选择
        pending_hr = getattr(self, "_current_header_row", 1)
        if self.pending_config:
            try:
                pending_hr = int(
                    self.pending_config["profile"]["excel"].get("header_row", 1)
                )
            except (KeyError, TypeError, ValueError):
                pending_hr = 1
        self._apply_header_row(pending_hr)
        self.current_excel_path = worker.file_path

        if result.get("source") == "wps_cloud":
            self.data_source = "wps_cloud"
            self.cloud_file_id = result.get("cloud_file_id", self.cloud_file_id)
            self.cloud_file_name = result.get(
                "cloud_file_name", self.cloud_file_name
            )
            # 全部 sheet 留在内存，切 sheet 不重新下载
            self._cloud_sheets_cache = result.get("all_sheets_data") or {}
            display_name = self.cloud_file_name or self.cloud_file_id
            self.current_file_label.setText(f"☁ 云文档: {display_name}")
        else:
            self.data_source = "local"
            self._cloud_sheets_cache = {}
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
            if self.data_source == "wps_cloud":
                self.log("云文档读取完成（临时文件已自动清理）！请选择Sheet和列，然后点击'加载数据'")
            else:
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
        if index < 0 or not self.current_excel_path:
            return
        sheet_name = self.sheet_combo.itemText(index)

        # 云文档：全部 sheet 已在内存，直接切换，不重新下载
        if self.data_source == "wps_cloud":
            values = self._cloud_sheets_cache.get(sheet_name)
            if values is None:
                self.log(f"云文档缓存中没有 Sheet: {sheet_name}，请点“重新读取”")
                return
            self.log(f"--- 切换到云文档Sheet: {sheet_name}（内存缓存）---")
            self._switch_sheet_in_memory(sheet_name, values)
            return

        self.log(f"--- 切换到Sheet: {sheet_name} ---")

        if self.excel_worker and self.excel_worker.isRunning():
            self.log("Excel文件正在读取，请稍候再切换Sheet")
            return

        self._load_excel_data(self.current_excel_path, sheet_name)

    def _switch_sheet_in_memory(self, sheet_name, values):
        """云文档切 Sheet：复用已下载到内存的数据（临时文件已删除）。"""
        self.current_sheet = sheet_name
        self.table_data = values
        self.headers = TableProcessor.derive_headers(values[0])
        self._apply_header_row(self._current_header_row)
        self.processor = None
        self.wechat_mapping = {}
        self.persons_list.clear()
        self.preview_text.clear()
        self.send_btn.setEnabled(False)
        self.start_send_btn.setEnabled(False)
        self.last_failed_tasks = []
        self.retry_send_btn.setEnabled(False)
        self.clear_filter_conditions()
        self._update_column_combos()

    def _apply_header_row(self, hr):
        """按表头行号(1-based)设置 spin 并据该行重建 self.headers（静默，不重置选择）。"""
        data = self.table_data or []
        if not data:
            return
        try:
            hr = int(hr)
        except (TypeError, ValueError):
            hr = 1
        hr = max(1, min(hr, len(data)))
        self.header_row_spin.blockSignals(True)
        try:
            self.header_row_spin.setMaximum(max(1, len(data)))
            self.header_row_spin.setValue(hr)
        finally:
            self.header_row_spin.blockSignals(False)
        self.headers = TableProcessor.derive_headers(data[hr - 1])
        self._current_header_row = hr

    def on_header_row_changed(self, hr):
        """用户手动改表头行：列结构随之变化，重置列选择与已加载数据。"""
        if not self.table_data:
            return
        if self.excel_worker and self.excel_worker.isRunning():
            return
        # 发送进行中改动表头行会导致进行中的任务数据错乱，直接回退
        if self.worker and self.worker.isRunning():
            QMessageBox.warning(self, "警告", "发送进行中不能修改表头行")
            self.header_row_spin.blockSignals(True)
            self.header_row_spin.setValue(getattr(self, "_current_header_row", 1))
            self.header_row_spin.blockSignals(False)
            return

        hr = max(1, min(int(hr), len(self.table_data)))
        if hr != self.header_row_spin.value():
            self.header_row_spin.blockSignals(True)
            self.header_row_spin.setValue(hr)
            self.header_row_spin.blockSignals(False)
        self.headers = TableProcessor.derive_headers(self.table_data[hr - 1])
        self._current_header_row = hr

        self._update_column_combos()
        self.processor = None
        self.wechat_mapping = {}
        self.persons_list.clear()
        self.preview_text.clear()
        self.send_btn.setEnabled(False)
        self.start_send_btn.setEnabled(False)
        self.last_failed_tasks = []
        self.retry_send_btn.setEnabled(False)
        self.clear_filter_conditions()
        self.log(f"表头行已切换为第 {hr} 行，请重新选择列并点击“加载数据”")

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
        
        self.processor = TableProcessor(
            self.table_data,
            header_row=self.header_row_spin.value(),
        )
        self.refresh_persons_list()
        
        if wechat_column:
            try:
                self.wechat_mapping = self.processor.get_person_to_wechat_mapping(name_column, wechat_column)
                self.log(f"已建立 {len(self.wechat_mapping)} 个微信映射")
            except Exception as e:
                self.log(f"⚠ 建立微信映射失败: {e}")
        
        self.log(f"数据加载完成！找到 {self.persons_list.count()} 个人")
        # 自动折叠左侧配置区，让人员列表和发送区更宽敞
        self._collapse_groups(self.config_group, self.url_group, self.filter_group)
        # 自动展开人员列表/数据预览/发送进度/发送控制。
        # 无论手动点“加载数据”还是加载配置后自动调用，都收敛到这一个入口，
        # 保证两条路径加载后的视图状态一致。
        for g in getattr(self, "_load_expand_groups", []):
            if g is not None and g.is_collapsed():
                g.set_collapsed(False)

    def log(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())
        
        # 用 QTextDocument 原生上限做裁剪，CPU 比手动逐块删除更低
        if not hasattr(self, "_log_max_lines_applied"):
            self.log_text.document().setMaximumBlockCount(1200)
            self._log_max_lines_applied = True

    def send_data(self, auto=False):
        # 并发守卫：手动发送/重试/auto_send 共用同一个微信单例，
        # 两个 worker 同时跑会争抢微信窗口、把消息发错对象。
        if self.worker is not None:
            try:
                busy = self.worker.isRunning()
            except Exception:
                busy = False
            if busy:
                if auto:
                    self.log("⚡ 自动发送：已有发送任务进行中，本次跳过（避免并发操作微信）")
                else:
                    QMessageBox.information(self, "提示", "已有发送任务进行中，请等待完成")
                return
        selected_items = self.persons_list.selectedItems()
        if not selected_items:
            if auto:
                self.log("自动发送：筛选结果为空，未发送")
                return
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
        send_order = self._get_send_order()
        attachment_path = self.attachment_edit.text().strip()

        # 至少要勾选一种实际会发送的内容
        effective_order = [k for k in send_order if not (
            (k == "custom" and not custom_msg)
            or (k == "attachment" and not attachment_path)
        )]
        if not effective_order:
            QMessageBox.warning(self, "警告", "请至少勾选一种要发送的内容（表格文字/表格图片/自定义消息/附加文件）")
            return
        
        tasks = []
        missing_wechat = []

        # 联系人映射：统一解析（优先 UI 面板，面板缺失时从当前/最近配置恢复）
        (_enabled, value_to_recipients, default_recipient, mapping_file,
         _from_ui) = self._resolve_mapping_for_send()
        mapping_enabled = _enabled
        # 未启用但已有映射条目/关联文件时醒目标志提醒，
        # 避免用户以为映射已生效（需勾选「启用联系人映射」）
        if not mapping_enabled and (
            value_to_recipients or mapping_file
        ):
            self.log(
                "⚠️ 检测到联系人映射条目/关联文件，但「启用联系人映射」未勾选，"
                "本次发送未应用映射，将按默认接收人逻辑发送（如需映射请勾选该开关）"
            )
        # 关联了映射表文件 → 发送前自动读取最新内容，
        # 用户改了文件不用手动"导入表格"
        if mapping_enabled:
            if mapping_file and os.path.isfile(mapping_file):
                try:
                    from modules.profile_runner import _load_mappings_from_file
                    fresh = _load_mappings_from_file(mapping_file, self.log)
                    if fresh:
                        fresh_map = {
                            m["source_value"]: list(m["recipients"])
                            for m in fresh
                        }
                        # 覆盖式合并：文件覆盖同名 key，UI 中有而文件没有的保留
                        value_to_recipients.update(fresh_map)
                        self.log(
                            f"已从关联文件自动加载 {len(fresh_map)} 条映射: {mapping_file}"
                        )
                except Exception as exc:
                    self.log(f"⚠ 自动读取映射表失败，使用 UI 当前数据: {exc}")

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

            # 接收人解析：映射优先 → 兜底 → 原"筛选列"逻辑
            fallback = (
                self.wechat_mapping.get(name, "")
                or self.wechat_edit.text().strip()
                or name
            )
            recipients_list = self._resolve_recipients_for_person(
                name, value_to_recipients, default_recipient,
                mapping_enabled, fallback,
            )
            if not recipients_list:
                # 全部兜底失败（仅当映射启用但未命中且无 fallback 时可能）
                missing_wechat.append(name)
                continue

            # 命中映射时明确提示"谁 → 映射给谁发送"，便于核对哪个名字被替换
            if mapping_enabled and name in value_to_recipients and len(recipients_list) > 0:
                self.log(f"◈ {name} → 映射给 {', '.join(recipients_list)} 发送")
            elif name != recipients_list[0]:
                self.log(f"◈ {name} → 发送给 {', '.join(recipients_list)}")

            for recipient in recipients_list:
                tasks.append({
                    "name": name,
                    "person_data": person_data,
                    "table_data": table_data,
                    "recipient": recipient,
                    "custom_msg": custom_msg,
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
        _order_label = {"text": "表格文字", "image": "表格图片", "custom": "自定义消息", "attachment": "附加文件"}
        confirm_text += "\n发送顺序: " + " → ".join(_order_label.get(k, k) for k in effective_order)
        if custom_msg:
            confirm_text += f"\n包含自定义消息: {custom_msg[:30]}..." if len(custom_msg) > 30 else f"\n包含自定义消息: {custom_msg}"
        
        if auto:
            # 加载配置后自动发送：已按配置全选，跳过人工确认直接发送
            self.log(f"⚡ 自动发送：已全选筛选出的 {len(tasks)} 人，开始发送")
        else:
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
        self.progress_bar.setRange(0, len(tasks))
        self.progress_label.setText("正在发送...")
        self.last_send_order = list(send_order)

        worker = SendWorker(
            tasks=tasks,
            send_interval=self.send_interval_spin.value(),
            chat_delay=self.chat_delay_spin.value(),
            send_order=list(send_order),
            minimize_after=self.minimize_check.isChecked(),
            attachment=self.attachment_edit.text().strip()
        )
        self.worker = worker
        # 记录本轮是否为自动发送（加载配置后无人值守）：完成后不弹模态确认框
        self._current_send_auto = bool(auto)
        if self.worker_created_cb:
            try:
                self.worker_created_cb(worker)
            except Exception:
                pass
        worker.signals.result.connect(self.on_send_result)
        worker.signals.error.connect(self.on_send_error)
        worker.signals.log.connect(self.log)
        worker.signals.progress.connect(self.on_send_progress)
        worker.finished.connect(
            lambda current=worker: self.on_send_finished(current)
        )
        worker.start()
        # 自动切换视图：只保留发送控制+发送进度，其他面板收起
        self._enter_sending_view()

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
        self.progress_bar.setRange(0, len(failed_tasks))
        self.progress_label.setText("正在重试发送...")
        
        worker = SendWorker(
            tasks=failed_tasks,
            send_interval=self.send_interval_spin.value(),
            chat_delay=self.chat_delay_spin.value(),
            send_order=list(getattr(self, 'last_send_order', ['text'])),
            minimize_after=self.minimize_check.isChecked(),
            attachment=self.attachment_edit.text().strip()
        )
        self.worker = worker
        # 重试均为用户手动操作，完成后正常弹确认框
        self._current_send_auto = False
        if self.worker_created_cb:
            try:
                self.worker_created_cb(worker)
            except Exception:
                pass
        worker.signals.result.connect(self.on_send_result)
        worker.signals.error.connect(self.on_send_error)
        worker.signals.log.connect(self.log)
        worker.signals.progress.connect(self.on_send_progress)
        worker.finished.connect(
            lambda current=worker: self.on_send_finished(current)
        )
        worker.start()
        self._enter_sending_view()

    def on_send_finished(self, worker):
        if worker is self.worker:
            self.worker = None
            self.send_btn.setEnabled(True)
            self.start_send_btn.setEnabled(True)
            self.pause_send_btn.setEnabled(False)
            self.stop_send_btn.setEnabled(False)
            self.pause_send_btn.setText("⏸ 暂停")
        worker.deleteLater()

    def on_send_progress(self, current, total):
        # 进度信号为 (当前人数, 总人数)；进度条范围在发送开始时已按总人数设置
        total = max(1, int(total))
        current = max(0, min(int(current), total))
        self.progress_bar.setValue(current)
        pct = int(current * 100 / total)
        self.progress_label.setText(f"发送进度: {current}/{total} ({pct}%)")

    def on_send_result(self, result):
        success_count, failed_count, total_count, failed_tasks = result
        self.last_failed_tasks = failed_tasks
        mw = self.window()

        # 自动发送（加载配置后无人值守）：不弹模态确认框，否则弹窗无人关闭，
        # 窗口标题/任务栏会一直定格在发送态。改用 Toast 通知 + 日志，并立即收尾。
        if getattr(self, "_current_send_auto", False):
            tip = f"自动发送完成：成功 {success_count}，失败 {failed_count}，总计 {total_count}"
            self.log(f"⚡ {tip}")
            try:
                ToastNotification.show_toast(
                    "表格自动发送", tip,
                    success=(failed_count == 0),
                    duration=8000, parent=mw)
            except Exception:
                pass
            # 无条件复位门闩并恢复标题/任务栏（失败时内部红色停留 4 秒）
            if self.progress_dismissed_cb:
                self.progress_dismissed_cb(failed_count)
            self.progress_label.setText(f"自动发送完成: 成功 {success_count}, 失败 {failed_count}")
            if failed_tasks:
                self.retry_send_btn.setEnabled(True)
                self.log(f"⚠ 有 {len(failed_tasks)} 个发送失败，可点击'重试发送'按钮重新发送")
                # auto_send 有失败时也自动展开日志面板
                mw2 = self.window()
                if mw2 is not None and hasattr(mw2, "_expand_log_on_failure"):
                    try:
                        mw2._expand_log_on_failure()
                    except Exception:
                        pass
            else:
                self.retry_send_btn.setEnabled(False)
            return

        # 手动发送完成：用 Toast 通知替代模态弹窗，避免无人值守时弹窗无人关闭
        # 导致标题/任务栏永久卡在发送态。Toast 自动消失且不阻塞。
        try:
            ToastNotification.show_toast(
                "发送完成",
                f"成功 {success_count}，失败 {failed_count}，总计 {total_count}",
                success=(failed_count == 0),
                duration=5000, parent=mw)
        except Exception:
            pass
        finally:
            # try/finally 兜底：无论 Toast 是否成功弹出，都必须复位门闩
            if self.progress_dismissed_cb:
                self.progress_dismissed_cb(failed_count)
        # 用户点掉完成弹窗后，再恢复窗口标题并清除任务栏定格进度
        self.progress_label.setText(f"发送完成: 成功 {success_count}, 失败 {failed_count}")

        if failed_tasks:
            self.retry_send_btn.setEnabled(True)
            self.log(f"⚠ 有 {len(failed_tasks)} 个发送失败，可点击'重试发送'按钮重新发送")
            # 手动发送有失败时也自动展开日志面板
            if mw is not None and hasattr(mw, "_expand_log_on_failure"):
                try:
                    mw._expand_log_on_failure()
                except Exception:
                    pass
        else:
            self.retry_send_btn.setEnabled(False)

    def on_send_error(self, error):
        mw = self.window()
        try:
            ToastNotification.show_toast(
                "发送出错", f"发送失败: {error}",
                success=False, duration=8000, parent=mw)
        except Exception:
            pass
        finally:
            # 同 on_send_result：Toast 异常时也必须复位门闩、恢复任务栏与标题
            if self.progress_dismissed_cb:
                self.progress_dismissed_cb(1)
        self.log(f"发送失败: {error}")
        self.progress_label.setText("发送出错")
        # 发送出错也自动展开日志面板
        if mw is not None and hasattr(mw, "_expand_log_on_failure"):
            try:
                mw._expand_log_on_failure()
            except Exception:
                pass


class ToastNotification(QFrame):
    """右下角弹出的通知窗口（类似 Windows Toast / 其他软件通知）。

    无边框、80% 不透明圆角、跟随系统亮/暗配色、自动消失、点击可关闭。
    主线程创建和操作，所有信号都走主线程事件循环，线程安全。
    失败通知不自动隐藏，需手动关闭。
    """

    # 类属性：当前显示中的通知实例（用列表支持多个通知叠加）
    _active_instances: list = []

    @classmethod
    def _is_dark_mode(cls) -> bool:
        """检测系统当前是否暗色模式（通知是瞬时弹出的，无需监听切换）。"""
        try:
            return QGuiApplication.styleHints().colorScheme() == \
                Qt.ColorScheme.Dark
        except Exception:
            return False

    @classmethod
    def show_toast(cls, title: str, message: str, *,
                    success: bool = True, duration: int = 8000,
                    parent=None, on_click=None):
        """弹出通知窗口。

        Args:
            title: 通知标题
            message: 通知内容
            success: True=绿色成功图标，False=红色错误图标
            duration: 成功通知自动消失毫秒数；失败通知强制不自动隐藏
            parent: 父 QWidget（点击通知时把它带到前台，通常传 MainWindow）
            on_click: 可选点击回调；提供时点击通知执行回调（由回调负责
                关闭通知和后续动作），不提供则保持默认“带到前台并关闭”
        """
        try:
            # 有发送失败的通知常驻，必须手动关闭
            if not success:
                duration = 0
            toast = cls(title, message, success=success,
                        duration=duration, parent=parent, on_click=on_click)
            cls._active_instances.append(toast)
            toast.destroyed.connect(
                lambda *_: cls._active_instances.remove(toast)
                if toast in cls._active_instances else None)
            toast.show()
            toast.raise_()
            # 注意：主窗口最小化到托盘时本进程没有前台窗口，Windows 前台锁定
            # 会让 activateWindow() 失效甚至使窗口停留在不可见的"鬼影"状态。
            # Toast 已带 WindowDoesNotAcceptFocus（WS_EX_NOACTIVATE），无需
            # 也不能抢焦点，因此这里不再调用 activateWindow()。
            # show 之后再用原生 API 强制置顶（后台进程首次 show 时 Qt 的
            # WindowStaysOnTopHint 可能被 Windows 忽略）。
            try:
                import ctypes
                hwnd = int(toast.winId())
                # HWND_TOPMOST=-1；SWP_NOMOVE|NOSIZE|NOACTIVATE|SHOWWINDOW
                ctypes.windll.user32.SetWindowPos(
                    hwnd, -1, 0, 0, 0, 0, 0x0002 | 0x0001 | 0x0010 | 0x0040)
            except Exception:
                pass
            # 入场动画
            toast._animate_in()
            try:
                logger.info(
                    "[Toast] 显示通知 title=%s success=%s 位置=(%s,%s) 尺寸=%sx%s 可见=%s",
                    title, success, toast.x(), toast.y(),
                    toast.width(), toast.height(), toast.isVisible())
            except Exception:
                pass
        except Exception:
            # 不再静默吞掉：记录异常便于定位"通知不显示"类问题
            logger.exception("[Toast] show_toast 失败 title=%s", title)

    def __init__(self, title: str, message: str, *,
                 success: bool = True, duration: int = 8000,
                 parent=None, on_click=None):
        # parent 用于点击时把主窗口带到前台；不强设为 Qt 父对象，避免窗口嵌入
        self._main_window = parent
        # 自定义点击回调（如“发现新版本”点击后打开更新对话框）
        self._on_click = on_click
        super().__init__(None)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Tool |
            Qt.WindowType.WindowStaysOnTopHint |
            # WS_EX_NOACTIVATE：主窗口隐藏到托盘、本进程无前台窗口时，
            # Windows 前台锁定会阻止通知窗口显现；不接受焦点的窗口不受此
            # 限制，可直接在右下角显示，且仍能正常接收鼠标点击。
            Qt.WindowType.WindowDoesNotAcceptFocus
        )
        # WA_StyledBackground 确保 stylesheet 背景色在打包后也生效
        # （WA_TranslucentBackground 在 Nuitka 打包后可能导致背景全透，弃用）
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self._duration = duration
        self._closing = False
        self._dark = self._is_dark_mode()

        # 跟随系统亮/暗模式的配色（完全不透明，确保文字清晰可读）
        if self._dark:
            bg_color = "rgb(37, 38, 41)"
            border_color = "rgba(255, 255, 255, 0.14)"
            title_color = "#f0f0f0"
            msg_color = "#c9ccd1"
            close_color = "#9aa0a6"
        else:
            bg_color = "rgb(255, 255, 255)"
            border_color = "rgba(0, 0, 0, 0.10)"
            title_color = "#2c3e50"
            msg_color = "#34495e"
            close_color = "#95a5a6"

        # 布局
        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(12)

        # 图标
        icon_label = QLabel()
        icon_label.setFixedSize(32, 32)
        if success:
            icon_label.setStyleSheet(
                "QLabel { background: #27ae60; border-radius: 16px; "
                "color: white; font-size: 18px; font-weight: bold; }")
            icon_label.setText("✓")
            icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        else:
            icon_label.setStyleSheet(
                "QLabel { background: #e74c3c; border-radius: 16px; "
                "color: white; font-size: 18px; font-weight: bold; }")
            icon_label.setText("✕")
            icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # 文字区
        text_layout = QVBoxLayout()
        text_layout.setSpacing(4)
        title_label = QLabel(title)
        title_label.setStyleSheet(
            f"color: {title_color}; font-size: 14px; "
            f"font-weight: bold; background: transparent;")
        msg_label = QLabel(message)
        msg_label.setStyleSheet(
            f"color: {msg_color}; font-size: 12px; "
            f"background: transparent;")
        msg_label.setWordWrap(True)
        msg_label.setMaximumWidth(280)
        text_layout.addWidget(title_label)
        text_layout.addWidget(msg_label)

        layout.addWidget(icon_label)
        layout.addLayout(text_layout, 1)

        # 关闭按钮（仅关闭，不带主窗口到前台）
        close_btn = QPushButton("✕")
        close_btn.setFixedSize(20, 20)
        close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        close_btn.setStyleSheet(
            f"QPushButton {{ border: none; color: {close_color}; "
            f"font-size: 14px; background: transparent; }}"
            "QPushButton:hover { color: #e74c3c; }")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn, 0, Qt.AlignmentFlag.AlignTop)

        # 整体样式
        self.setStyleSheet(
            f"ToastNotification {{ "
            f"background: {bg_color}; "
            f"border-radius: 10px; "
            f"border: 1px solid {border_color}; }}")
        # 鼠标手型，提示可点击
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedSize(380, max(70, 30 + 20 * (message.count('\n') + 1)))

        # 非按钮子 widget 鼠标事件穿透，使点击通知任意区域都触发 QFrame.mousePressEvent
        for w in (icon_label, title_label, msg_label):
            w.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

        # 定位到右下角（叠加排列）
        screen = QApplication.primaryScreen()
        if screen is not None:
            geo = screen.availableGeometry()
        else:
            geo = self.geometry()
        # 从右下角往上叠加
        offset_y = 10
        for prev in ToastNotification._active_instances:
            if prev is not self and prev.isVisible():
                offset_y += prev.height() + 8
        x = geo.x() + geo.width() - self.width() - 16
        y = geo.y() + geo.height() - self.height() - offset_y
        self.move(x, y)

        # 自动消失定时器（失败通知 duration=0，常驻直到手动关闭）
        if duration > 0:
            QTimer.singleShot(duration, self._fade_out)

    def _find_main_window(self):
        """找到主窗口：优先 parent，其次顶层 MainWindow。"""
        mw = self._main_window
        if mw is not None:
            return mw
        for w in QApplication.topLevelWidgets():
            if w.isWindow() and w.__class__.__name__ == "MainWindow":
                return w
        return None

    def _bring_main_to_front(self):
        """把主窗口从托盘/最小化恢复到前台。"""
        try:
            mw = self._find_main_window()
            if mw is not None:
                mw.showNormal()
                mw.raise_()
                mw.activateWindow()
        except Exception:
            pass

    def _animate_in(self):
        """入场动画：从右侧滑入 + 渐显。"""
        self._anim = QPropertyAnimation(self, b"pos")
        start = self.pos()
        end = QPoint(start.x() + 50, start.y())
        # 渐显
        self.setWindowOpacity(0.0)
        self._fade_anim = QPropertyAnimation(self, b"windowOpacity")
        self._fade_anim.setDuration(300)
        self._fade_anim.setStartValue(0.0)
        self._fade_anim.setEndValue(1.0)
        self._fade_anim.start()
        self._anim.setDuration(300)
        self._anim.setStartValue(end)
        self._anim.setEndValue(start)
        self._anim.start()

    def _fade_out(self):
        """退场动画：渐隐后关闭。"""
        if self._closing:
            return
        self._closing = True
        self._fade_out_anim = QPropertyAnimation(self, b"windowOpacity")
        self._fade_out_anim.setDuration(300)
        self._fade_out_anim.setStartValue(1.0)
        self._fade_out_anim.setEndValue(0.0)
        self._fade_out_anim.finished.connect(self.close)
        self._fade_out_anim.start()

    def mousePressEvent(self, event):
        """点击通知区域：有自定义回调走回调；默认把主窗口带到前台并关闭。"""
        if event.button() == Qt.MouseButton.LeftButton:
            if self._on_click is not None:
                try:
                    self._on_click()
                except Exception:
                    logger.exception("Toast 点击回调执行失败")
                return
            self._bring_main_to_front()
            self.close()


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
        # MainWindow 注入：worker 创建后回调，用于连接任务栏进度等全局信号
        self.worker_created_cb = None
        # 表单脏检测：_loading_form 期间程序化填充表单不视为改动；
        # _form_ready 在首次任务加载完成前为 False（避免启动时误弹保存区）
        self._loading_form = False
        self._form_dirty = False
        self._form_ready = False
        self.init_ui()
        self.connect_signals()
        self.reload_tasks()
        self._form_ready = True

    # ----------------------------- UI -----------------------------
    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(8)

        # --- 左侧：任务列表 + 配置保存（芯片排，可同开；保存区默认收起，
        #     表单有改动/新建任务时自动果冻弹出，保存成功后自动收起） ---
        left_panel = QWidget()
        left_column = QVBoxLayout(left_panel)
        left_column.setContentsMargins(0, 0, 0, 0)
        left_column.setSpacing(6)
        left_bar = ChipBar(exclusive=False)
        left_column.addWidget(left_bar)

        list_group = left_bar.add_section("📋 任务列表")
        left_layout = list_group.contentLayout()

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

        left_column.addWidget(list_group)

        save_group = left_bar.add_section("💾 配置保存", collapsed=True)
        save_layout = save_group.contentLayout()
        left_btn2 = QHBoxLayout()
        self.save_all_btn = QPushButton("保存为定时配置")
        self.load_config_btn = QPushButton("加载定时配置")
        left_btn2.addWidget(self.save_all_btn)
        left_btn2.addWidget(self.load_config_btn)
        save_layout.addLayout(left_btn2)
        left_column.addWidget(save_group)

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
        main_layout.addWidget(left_panel, 0)

        # --- 右侧：编辑表单（手风琴芯片排：同一时间只展开一个面板） ---
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)
        right_bar = ChipBar(exclusive=True)
        right_layout.addWidget(right_bar)

        # 基本
        base_group = right_bar.add_section("📌 任务基础")
        base_layout = base_group.contentLayout()
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
        repeat_group = right_bar.add_section("🔁 重复规则", collapsed=True)
        repeat_layout = repeat_group.contentLayout()
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("模式:"))
        self.repeat_mode_combo = QComboBox()
        self.repeat_mode_combo.addItem("每天", "daily")
        self.repeat_mode_combo.addItem("每周指定日期", "weekly")
        self.repeat_mode_combo.addItem("指定日期执行(一次)", "once")
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

        self.date_group_box = QGroupBox("执行日期(一次性模式生效，可添加多个):")
        date_layout = QVBoxLayout(self.date_group_box)
        self.date_list = QListWidget()
        self.date_list.setMaximumHeight(72)
        date_layout.addWidget(self.date_list)
        date_row = QHBoxLayout()
        self.run_date_edit = QDateEdit()
        self.run_date_edit.setCalendarPopup(True)
        self.run_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.run_date_edit.setDate(QDate.currentDate())
        date_row.addWidget(self.run_date_edit)
        self.add_date_btn = QPushButton("添加")
        self.del_date_btn = QPushButton("删除选中")
        date_row.addWidget(self.add_date_btn)
        date_row.addWidget(self.del_date_btn)
        date_layout.addLayout(date_row)
        repeat_layout.addWidget(self.date_group_box)
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
        send_group = right_bar.add_section("✉ 发送内容", collapsed=True)
        send_layout = send_group.contentLayout()

        # 任务类型：纯文字消息 / 执行表格配置（可多个链式执行）
        kind_row = QHBoxLayout()
        kind_row.addWidget(QLabel("任务类型:"))
        self.kind_combo = QComboBox()
        self.kind_combo.addItem("发送文字消息", "message")
        self.kind_combo.addItem(
            "执行表格配置（可多个配置链式执行）", "profiles"
        )
        kind_row.addWidget(self.kind_combo, 1)
        send_layout.addLayout(kind_row)

        # profiles 模式：配置文件列表 + 添加/移除/上移/下移
        self.profile_panel = QWidget()
        profile_panel_layout = QVBoxLayout(self.profile_panel)
        profile_panel_layout.setContentsMargins(0, 0, 0, 0)
        profile_panel_layout.addWidget(
            QLabel("按顺序执行的配置文件（执行完上一个再执行下一个）:")
        )
        self.profile_list = QListWidget()
        self.profile_list.setMinimumHeight(120)
        profile_panel_layout.addWidget(self.profile_list)
        profile_btn_row = QHBoxLayout()
        self.profile_add_btn = QPushButton("➕ 添加配置")
        self.profile_remove_btn = QPushButton("➖ 移除选中")
        self.profile_up_btn = QPushButton("⬆ 上移")
        self.profile_down_btn = QPushButton("⬇ 下移")
        profile_btn_row.addWidget(self.profile_add_btn)
        profile_btn_row.addWidget(self.profile_remove_btn)
        profile_btn_row.addWidget(self.profile_up_btn)
        profile_btn_row.addWidget(self.profile_down_btn)
        profile_btn_row.addStretch()
        profile_panel_layout.addLayout(profile_btn_row)
        profile_hint = QLabel(
            "提示：配置来自「数据发送」页保存的配置文件，\n"
            "本地表格和云文档配置都支持；接收人/筛选/发送顺序\n"
            "均以各配置自己的设置为准。"
        )
        profile_hint.setStyleSheet("color:#888; font-size:11px;")
        profile_panel_layout.addWidget(profile_hint)
        send_layout.addWidget(self.profile_panel)

        # message 模式面板（原接收人/消息/附件/顺序/延迟设置）
        self.message_panel = QWidget()
        msg_layout = QVBoxLayout(self.message_panel)
        msg_layout.setContentsMargins(0, 0, 0, 0)

        msg_layout.addWidget(QLabel("接收人(好友/群名，支持模糊匹配，每行一个或英文逗号分隔):"))
        self.recipients_edit = QTextEdit()
        self.recipients_edit.setMaximumHeight(70)
        self.recipients_edit.setPlaceholderText("张三\n文件传输助手\n工作群A")
        msg_layout.addWidget(self.recipients_edit)

        msg_layout.addWidget(QLabel("自定义文字消息:"))
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
        msg_layout.addWidget(self.message_edit)

        city_row = QHBoxLayout()
        city_row.addWidget(QLabel("默认城市(天气占位符不带参时用此):"))
        self.default_city_edit = QLineEdit()
        self.default_city_edit.setPlaceholderText("如：北京；留空则用「天气设置」里的全局默认城市")
        city_row.addWidget(self.default_city_edit, 1)
        msg_layout.addLayout(city_row)

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
        msg_layout.addLayout(attach_row)

        order_caption = QLabel("发送顺序(勾选要发送的内容，选中后点右侧按钮调整先后):")
        msg_layout.addWidget(order_caption)
        sched_order_row = QHBoxLayout()
        self.sched_order_list = QListWidget()
        self.sched_order_list.setMaximumHeight(54)
        self._SCHED_ORDER_ITEMS = [
            ("message", "消息文字"),
            ("attachment", "附加文件"),
        ]
        for key, label in self._SCHED_ORDER_ITEMS:
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, key)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(
                Qt.CheckState.Checked if key == "message" else Qt.CheckState.Unchecked
            )
            self.sched_order_list.addItem(item)
        sched_order_row.addWidget(self.sched_order_list, 1)
        sched_order_btns = QVBoxLayout()
        self.sched_order_up_btn = QPushButton("⬆ 上移")
        self.sched_order_down_btn = QPushButton("⬇ 下移")
        sched_order_btns.addWidget(self.sched_order_up_btn)
        sched_order_btns.addWidget(self.sched_order_down_btn)
        sched_order_row.addLayout(sched_order_btns)
        msg_layout.addLayout(sched_order_row)

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
        msg_layout.addLayout(delay_row)

        send_layout.addWidget(self.message_panel)

        right_layout.addWidget(send_group)

        # 电脑锁定/完成后行为配置（任务级，每个任务可单独设置）
        lock_group = right_bar.add_section("🔒 锁定 / 完成后行为", collapsed=True)
        lock_layout = QHBoxLayout()
        lock_group.contentLayout().addLayout(lock_layout)
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

        # 操作区（默认收起：表单有改动/新建任务时自动弹出，也可手动点开）
        action_group = right_bar.add_section("💾 保存 / 重置", collapsed=True)
        action_row = QHBoxLayout()
        self.save_task_btn = QPushButton("💾 保存当前任务")
        self.save_task_btn.setStyleSheet(
            "background-color:#2196F3; color:white; padding:6px;"
        )
        self.reset_form_btn = QPushButton("重置表单")
        action_row.addWidget(self.save_task_btn)
        action_row.addWidget(self.reset_form_btn)
        action_row.addStretch()
        action_group.contentLayout().addLayout(action_row)
        right_layout.addWidget(action_group)

        # 日志（共享主日志的回调，这里也放只读面板，方便查看）
        log_group = right_bar.add_section("📜 定时发送日志", collapsed=True)
        log_layout = log_group.contentLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(140)
        log_layout.addWidget(self.log_text)
        right_layout.addWidget(log_group)
        right_layout.addStretch(1)

        main_layout.addWidget(right, 1)

        # 注册所有可折叠面板，供全局精简开关遍历
        self._collapsible_groups = [
            list_group, save_group, base_group, repeat_group,
            send_group, lock_group, action_group, log_group,
        ]
        self.list_group = list_group
        self.save_group = save_group
        self.action_group = action_group
        self.schedule_log_group = log_group

    def _collapse_groups(self, *groups):
        """批量折叠指定分组（忽略 None）。"""
        for g in groups:
            if g is not None:
                g.set_collapsed(True)

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
        self.add_date_btn.clicked.connect(self._on_add_date)
        self.del_date_btn.clicked.connect(self._on_del_date)
        self.repeat_mode_combo.currentIndexChanged.connect(self._on_repeat_mode_changed)
        self.attach_pick_btn.clicked.connect(self._on_pick_attachment)
        self.attach_clear_btn.clicked.connect(self._on_clear_attachment)
        self.sched_order_up_btn.clicked.connect(lambda: self._move_sched_order_item(-1))
        self.sched_order_down_btn.clicked.connect(lambda: self._move_sched_order_item(1))
        self.sched_order_list.itemChanged.connect(self._on_sched_order_item_changed)
        self.kind_combo.currentIndexChanged.connect(self._on_kind_changed)
        self.profile_add_btn.clicked.connect(self._on_add_profiles)
        self.profile_remove_btn.clicked.connect(self._on_remove_profile)
        self.profile_up_btn.clicked.connect(lambda: self._move_profile_item(-1))
        self.profile_down_btn.clicked.connect(lambda: self._move_profile_item(1))
        self.save_all_btn.clicked.connect(self._on_save_all)
        self.load_config_btn.clicked.connect(self._on_load_all)
        self.run_now_btn.clicked.connect(self._on_run_now)
        self.stop_send_btn.clicked.connect(self._on_stop_send)
        self.weather_settings_btn.clicked.connect(self._on_open_weather_settings)
        self._connect_form_dirty_signals()

    # ------------------------- 表单脏检测 -------------------------
    def _connect_form_dirty_signals(self):
        """所有表单控件改动 → 标记脏并自动果冻弹出「配置保存」面板。"""
        mark = self._mark_form_dirty
        self.name_edit.textChanged.connect(mark)
        self.enabled_check.stateChanged.connect(mark)
        self.repeat_mode_combo.currentIndexChanged.connect(mark)
        for cb in self.weekday_checks.values():
            cb.stateChanged.connect(mark)
        self.time_list.model().rowsInserted.connect(lambda *a: mark())
        self.time_list.model().rowsRemoved.connect(lambda *a: mark())
        self.date_list.model().rowsInserted.connect(lambda *a: mark())
        self.date_list.model().rowsRemoved.connect(lambda *a: mark())
        self.recipients_edit.textChanged.connect(mark)
        self.message_edit.textChanged.connect(mark)
        self.default_city_edit.textChanged.connect(mark)
        self.chat_delay_spin.valueChanged.connect(mark)
        self.send_interval_spin.valueChanged.connect(mark)
        self.keep_unlock_check.stateChanged.connect(mark)
        self.relock_check.stateChanged.connect(mark)
        self.minimize_check.stateChanged.connect(mark)
        self.sched_order_list.itemChanged.connect(lambda *a: mark())
        self.kind_combo.currentIndexChanged.connect(mark)
        self.profile_list.model().rowsInserted.connect(lambda *a: mark())
        self.profile_list.model().rowsRemoved.connect(lambda *a: mark())

    def _mark_form_dirty(self, *_args):
        """表单被用户改动：置脏并自动弹出左侧配置保存区+右侧保存/重置操作区
        （程序填充阶段忽略）。"""
        if not self._form_ready or self._loading_form:
            return
        self._form_dirty = True
        if self.save_group.is_collapsed():
            self.save_group.set_collapsed(False, animate=True)
        if self.action_group.is_collapsed():
            self.action_group.set_collapsed(
                False, animate=True, accordion_close=False)

    def _clear_form_dirty(self):
        """保存/加载成功后清脏并自动收起保存区与操作区。"""
        self._form_dirty = False
        if not self.save_group.is_collapsed():
            self.save_group.set_collapsed(True, animate=True)
        if not self.action_group.is_collapsed():
            self.action_group.set_collapsed(True, animate=True)

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
        elif task.repeat_mode == "once":
            dates = getattr(task, "run_dates", None) or []
            rule = f"一次: {'、'.join(dates) if dates else '未指定日期'}"
        else:
            days = ",".join(WEEKDAY_NAMES[d - 1] for d in task.days) if task.days else "未选"
            rule = f"每周: {days}"
        times = "、".join(task.times) if task.times else "(无时间点)"
        if getattr(task, "kind", "message") == "profiles":
            target = f"配置链:{len(task.profile_paths or [])}个"
        else:
            target = f"人数:{len(task.recipients)}"
        return f"{mark} {task.name}  | {rule} | {times} | {target}"

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
        # 程序化填充：屏蔽脏信号，避免加载已有任务也弹出保存区
        self._loading_form = True
        try:
            self.name_edit.setText(task.name)
            self.enabled_check.setChecked(task.enabled)
            mode_index = {"daily": 0, "weekly": 1, "once": 2}.get(
                getattr(task, "repeat_mode", "daily"), 0
            )
            self.repeat_mode_combo.setCurrentIndex(mode_index)
            self.date_list.clear()
            for rd in (getattr(task, "run_dates", None) or []):
                self.date_list.addItem(rd)
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
            order = getattr(task, "send_order", None) or ["message"]
            # 附件为空时即使顺序里含 attachment 也不勾选
            if not self.attachment_edit.text().strip():
                order = [k for k in order if k != "attachment"]
                if "message" not in order:
                    order = ["message"] + order
            self._set_sched_order(order)
            kind = getattr(task, "kind", "message") or "message"
            kind_idx = 0
            for ki in range(self.kind_combo.count()):
                if self.kind_combo.itemData(ki) == kind:
                    kind_idx = ki
                    break
            self.kind_combo.setCurrentIndex(kind_idx)
            self._set_profile_paths(getattr(task, "profile_paths", None) or [])
            self._on_kind_changed()
            self._on_repeat_mode_changed(mode_index)
        finally:
            self._loading_form = False
        self._form_dirty = False
        # 加载已有任务 = 干净状态：自动收起保存区与操作区
        # （即时，避免连续切任务时动画打架）
        if not self.save_group.is_collapsed():
            self.save_group.set_collapsed(True)
        if not self.action_group.is_collapsed():
            self.action_group.set_collapsed(True)

    def _reset_form(self):
        self._loading_form = True
        try:
            self.current_task_id = None
            self.name_edit.clear()
            self.enabled_check.setChecked(True)
            self.repeat_mode_combo.setCurrentIndex(0)
            self.run_date_edit.setDate(QDate.currentDate())
            self.date_list.clear()
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
            self._set_sched_order(["message"])
            self.kind_combo.setCurrentIndex(0)
            self._set_profile_paths([])
            self._on_kind_changed()
        finally:
            self._loading_form = False
        self._form_dirty = False

    # ------------------------- 增删改 -------------------------
    def _on_new_task(self):
        self._reset_form()
        self.name_edit.setFocus()
        # 新任务 = 待保存的新配置：自动弹出左侧保存区+右侧操作区
        self._form_dirty = True
        if self.save_group.is_collapsed():
            self.save_group.set_collapsed(False, animate=True)
        if self.action_group.is_collapsed():
            self.action_group.set_collapsed(
                False, animate=True, accordion_close=False)

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
        # 保存成功：清脏并自动收起保存区
        self._clear_form_dirty()

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
        run_dates: List[str] = []
        if repeat_mode == "weekly":
            days = sorted([d for d, cb in self.weekday_checks.items() if cb.isChecked()])
            if not days:
                raise ValueError("每周模式请至少选择一个周几")
        elif repeat_mode == "once":
            from datetime import datetime as _dt
            for i in range(self.date_list.count()):
                d = self.date_list.item(i).text().strip()
                try:
                    _dt.strptime(d, "%Y-%m-%d")
                except ValueError:
                    continue
                if d not in run_dates:
                    run_dates.append(d)
            if not run_dates:
                raise ValueError("请至少添加一个执行日期")
        times: List[str] = []
        for i in range(self.time_list.count()):
            slot = _normalize_time(self.time_list.item(i).text())
            if slot and slot not in times:
                times.append(slot)
        if not times:
            raise ValueError("请至少添加一个发送时间点")
        kind = self.kind_combo.currentData() or "message"
        profile_paths = self._profile_paths_in_list()
        recipients = []
        message = ""
        attachment = ""
        send_order = ["message"]
        if kind == "profiles":
            if not profile_paths:
                raise ValueError("请至少添加一个要执行的配置文件")
        else:
            recipients = _parse_recipients(self.recipients_edit.toPlainText())
            if not recipients:
                raise ValueError("请至少填写一个接收人(好友名/群名)")
            message = self.message_edit.toPlainText()
            if not message.strip():
                raise ValueError("请填写自定义文字消息")
            attachment = self.attachment_edit.text().strip()
            send_order = self._get_sched_order()
        task = ScheduleTask(
            id=_new_task_id() if new_id else (self.current_task_id or _new_task_id()),
            name=name,
            enabled=self.enabled_check.isChecked(),
            repeat_mode=repeat_mode,
            days=days,
            run_dates=run_dates,
            times=times,
            recipients=recipients,
            message=message,
            chat_delay=float(self.chat_delay_spin.value()),
            send_interval=float(self.send_interval_spin.value()),
            default_city=self.default_city_edit.text().strip(),
            keep_unlocked=self.keep_unlock_check.isChecked(),
            relock_after=self.relock_check.isChecked(),
            minimize_after=self.minimize_check.isChecked(),
            attachment=attachment,
            send_order=send_order,
            kind=kind,
            profile_paths=profile_paths,
        )
        return task

    # ------------------------- 发送顺序 -------------------------
    def _get_sched_order(self):
        order = []
        for i in range(self.sched_order_list.count()):
            item = self.sched_order_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                order.append(item.data(Qt.ItemDataRole.UserRole))
        if "message" not in order:
            order.append("message")
        return order

    def _find_sched_order_item(self, key):
        for i in range(self.sched_order_list.count()):
            item = self.sched_order_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == key:
                return i, item
        return -1, None

    def _set_sched_item_checked(self, key, checked):
        _idx, item = self._find_sched_order_item(key)
        if item is None:
            return
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        if item.checkState() != state:
            self.sched_order_list.blockSignals(True)
            try:
                item.setCheckState(state)
            finally:
                self.sched_order_list.blockSignals(False)

    def _set_sched_order(self, order_keys):
        label_map = dict(self._SCHED_ORDER_ITEMS)
        self.sched_order_list.blockSignals(True)
        try:
            self.sched_order_list.clear()
            enabled = set(order_keys)
            ordered = list(order_keys)
            for key, _label in self._SCHED_ORDER_ITEMS:
                if key not in enabled and key not in ordered:
                    ordered.append(key)
            for key in ordered:
                item = QListWidgetItem(label_map.get(key, key))
                item.setData(Qt.ItemDataRole.UserRole, key)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(
                    Qt.CheckState.Checked if key in enabled else Qt.CheckState.Unchecked
                )
                self.sched_order_list.addItem(item)
        finally:
            self.sched_order_list.blockSignals(False)

    def _move_sched_order_item(self, delta):
        row = self.sched_order_list.currentRow()
        if row < 0:
            return
        target = row + delta
        if target < 0 or target >= self.sched_order_list.count():
            return
        item = self.sched_order_list.takeItem(row)
        self.sched_order_list.insertItem(target, item)
        self.sched_order_list.setCurrentRow(target)
        self._mark_form_dirty()

    def _on_sched_order_item_changed(self, _item):
        # 消息文字是定时任务的根本，若被取消勾选则提示并恢复勾选
        _idx, msg_item = self._find_sched_order_item("message")
        if msg_item is not None and msg_item.checkState() != Qt.CheckState.Checked:
            self.sched_order_list.blockSignals(True)
            try:
                msg_item.setCheckState(Qt.CheckState.Checked)
            finally:
                self.sched_order_list.blockSignals(False)
            self.log("[定时] 消息文字为必发项，已保持勾选")

    # ------------------------- 任务类型 / 配置链 -------------------------
    def _on_kind_changed(self, _index=None):
        is_profiles = self.kind_combo.currentData() == "profiles"
        self.profile_panel.setVisible(is_profiles)
        self.message_panel.setVisible(not is_profiles)

    def _profile_paths_in_list(self):
        paths = []
        for i in range(self.profile_list.count()):
            p = self.profile_list.item(i).data(Qt.ItemDataRole.UserRole)
            if p:
                paths.append(str(p))
        return paths

    def _on_add_profiles(self):
        try:
            from modules.config_manager import ConfigManager
            start_dir = str(ConfigManager().profile_dir)
            if not os.path.isdir(start_dir):
                start_dir = os.path.expanduser("~")
        except Exception:
            start_dir = os.path.expanduser("~")
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "选择一个或多个数据发送配置（可多选，按选择顺序链式执行）",
            start_dir,
            "数据发送配置 (*.json)",
        )
        if not paths:
            return
        existing = set(self._profile_paths_in_list())
        for p in paths:
            p = os.path.normpath(str(p))
            if p in existing:
                continue
            existing.add(p)
            item = QListWidgetItem(
                f"{self.profile_list.count() + 1}. {os.path.basename(p)}"
            )
            item.setData(Qt.ItemDataRole.UserRole, p)
            item.setToolTip(p)
            self.profile_list.addItem(item)
        self._mark_form_dirty()

    def _on_remove_profile(self):
        for item in list(self.profile_list.selectedItems()):
            self.profile_list.takeItem(self.profile_list.row(item))
        self._renumber_profile_items()
        self._mark_form_dirty()

    def _move_profile_item(self, delta):
        row = self.profile_list.currentRow()
        if row < 0:
            return
        target = row + delta
        if target < 0 or target >= self.profile_list.count():
            return
        item = self.profile_list.takeItem(row)
        self.profile_list.insertItem(target, item)
        self.profile_list.setCurrentRow(target)
        self._renumber_profile_items()
        self._mark_form_dirty()

    def _renumber_profile_items(self):
        for i in range(self.profile_list.count()):
            item = self.profile_list.item(i)
            p = item.data(Qt.ItemDataRole.UserRole) or ""
            item.setText(f"{i + 1}. {os.path.basename(str(p))}")

    def _set_profile_paths(self, paths):
        self.profile_list.clear()
        for p in paths or []:
            p = os.path.normpath(str(p))
            item = QListWidgetItem(
                f"{self.profile_list.count() + 1}. {os.path.basename(p)}"
            )
            item.setData(Qt.ItemDataRole.UserRole, p)
            item.setToolTip(p)
            self.profile_list.addItem(item)

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
            self._set_sched_item_checked("attachment", True)
            self._mark_form_dirty()

    def _on_clear_attachment(self):
        self.attachment_edit.clear()
        self._set_sched_item_checked("attachment", False)
        self._mark_form_dirty()

    # ------------------------- 重复模式/时间点 -------------------------
    def _on_repeat_mode_changed(self, index):
        mode = self.repeat_mode_combo.currentData()
        weekly = mode == "weekly"
        once = mode == "once"
        self.weekday_group_box.setEnabled(weekly)
        self.date_group_box.setVisible(once)

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

    def _on_add_date(self):
        d = self.run_date_edit.date().toString("yyyy-MM-dd")
        for i in range(self.date_list.count()):
            if self.date_list.item(i).text() == d:
                return
        self.date_list.addItem(d)

    def _on_del_date(self):
        for item in list(self.date_list.selectedItems()):
            self.date_list.takeItem(self.date_list.row(item))

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
        # 保存成功：清脏并自动收起保存区
        self._clear_form_dirty()

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
        # 加载完成 = 干净状态：收起保存区与操作区
        self._form_dirty = False
        if not self.save_group.is_collapsed():
            self.save_group.set_collapsed(True)
        if not self.action_group.is_collapsed():
            self.action_group.set_collapsed(True)

    # ------------------------- 立即执行 -------------------------
    def _on_run_now(self):
        if self.send_worker and self.send_worker.isRunning():
            QMessageBox.information(self, "提示", "已有定时发送正在进行中")
            return
        # 与手动表格发送/后台定时触发互斥：共用微信单例，不能并发驱动 wxauto
        mw = self.window()
        if mw is not None and hasattr(mw, "_is_other_send_running"):
            try:
                if mw._is_other_send_running(exclude=self.send_worker):
                    QMessageBox.information(
                        self,
                        "提示",
                        "微信正在执行其他发送任务（手动发送或定时任务），请等待完成后再执行",
                    )
                    return
            except Exception:
                pass
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

        # 关键：worker 在子线程里调用 log_callback，不能直接碰 QWidget。
        # 这里把日志投递切回主线程，避免跨线程访问控件导致的偶发崩溃/挂起。
        if getattr(current, "kind", "message") == "profiles":
            worker = ProfileChainWorker(
                profile_paths=list(current.profile_paths or []),
                log_callback=self._worker_log_signal.emit,
                minimize_after=getattr(current, "minimize_after", True),
            )
        else:
            worker = ScheduleSendWorker(
                recipients=current.recipients,
                message=current.message,
                chat_delay=current.chat_delay,
                send_interval=current.send_interval,
                log_callback=self._worker_log_signal.emit,
                default_city=current.default_city,
                minimize_after=getattr(current, "minimize_after", True),
                attachment=getattr(current, "attachment", "") or "",
                send_order=list(getattr(current, "send_order", None) or ["message"]),
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
        if self.worker_created_cb:
            try:
                self.worker_created_cb(worker)
            except Exception:
                pass
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
        stopped_any = False
        if self.send_worker and self.send_worker.isRunning():
            self.send_worker.stop()
            stopped_any = True
        # 调度器后台触发的 worker（可能是长时间运行的配置链）也要能停
        mw = self.window()
        main_worker = getattr(mw, "_schedule_worker", None) if mw is not None else None
        if main_worker is not None:
            try:
                if main_worker.isRunning() and hasattr(main_worker, "stop"):
                    main_worker.stop()
                    stopped_any = True
            except Exception:
                pass
        if stopped_any:
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
        # Toast 通知（立即执行完成）
        mw = self.window()
        try:
            ToastNotification.show_toast(
                "发送完成",
                f"成功 {success}，失败 {failed}",
                success=(failed == 0),
                duration=5000, parent=mw)
        except Exception:
            pass
        # 强制兜底恢复标题/任务栏：立即执行路径同样不依赖进度信号是否正确到达
        mw = self.window()
        if mw is not None and hasattr(mw, "_force_finish_schedule_progress"):
            try:
                mw._force_finish_schedule_progress(failed)
            except Exception:
                pass
        # 有失败时展开定时页日志面板（_force_finish_schedule_progress 内已有展开逻辑，
        # 但加在这里做双重保证，且展开自身面板时用 set_collapsed 更快）
        if failed and hasattr(self, "schedule_log_group"):
            try:
                lg = self.schedule_log_group
                if lg is not None and lg.is_collapsed():
                    lg.set_collapsed(False, animate=True)
            except Exception:
                pass
        # 按任务配置决定是否发送后锁定
        task = getattr(self, "_current_run_task", None)
        if task and getattr(task, "relock_after", False):
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

        # 测试连接状态：防重复点击 + 对话框关闭后回调安全处理
        test_state = {"running": False, "sig": None}

        def on_test():
            # 重入保护：测试进行中直接忽略后续点击（按钮也已置灰，双保险）
            if test_state["running"]:
                return
            if not dlg.isVisible():
                return
            test_city = city_edit.text().strip() or "北京"
            base_url = url_edit.text().strip()
            api_key = key_edit.text().strip()
            self.log(f"[天气] 开始测试连接：city={test_city} url={base_url}")
            # 防止重复点击 + 网络请求放后台线程，避免阻塞主线程导致界面卡死
            test_state["running"] = True
            test_btn.setEnabled(False)
            test_btn.setText("测试中...")
            result_label.setText("正在测试连接，请稍候（最多约 8 秒）...")

            class _TestSignal(QObject):
                done = pyqtSignal(bool, str)
                log = pyqtSignal(str)

            sig = _TestSignal()
            test_state["sig"] = sig

            def on_log(msg):
                try:
                    self.log(f"[天气] {msg}")
                except Exception:
                    pass

            def on_done(ok, m):
                # 回调到达时对话框可能已关闭，控件已销毁——整段保护
                try:
                    if dlg.isVisible():
                        color = "#2e7d32" if ok else "#c62828"
                        result_label.setText(f'<span style="color:{color}">{m}</span>')
                        result_label.setTextFormat(Qt.TextFormat.RichText)
                        test_btn.setEnabled(True)
                        test_btn.setText("测试连接")
                except Exception:
                    pass
                test_state["running"] = False
                test_state["sig"] = None
                try:
                    sig.deleteLater()
                except Exception:
                    pass

            sig.log.connect(on_log)
            sig.done.connect(on_done)

            def _worker():
                try:
                    ok, m = weather_fetcher.test_connection(
                        api_key,
                        base_url,
                        test_city,
                        log_fn=lambda msg: sig.log.emit(str(msg)),
                    )
                except Exception as exc:
                    ok, m = False, f"测试异常: {exc}"
                    try:
                        sig.log.emit(str(m))
                    except Exception:
                        pass
                try:
                    sig.done.emit(bool(ok), str(m))
                except Exception:
                    pass

            threading.Thread(target=_worker, daemon=True, name="WeatherTest").start()

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


class MonitorTab(QWidget):
    """📡 监控 标签页：本地目录监控 → 筛选/增量对比 → 推送微信。

    任务列表(左) + 编辑表单(右，芯片排面板)。持有一个常驻轮询 MonitorWorker，
    可手动"立即执行一次"。日志经信号投递回主线程写入面板与 app.log。
    """

    # worker 线程日志 → 信号投递回主线程再写面板（遵循 ScheduleTab 跨线程惯例）
    _worker_log_signal = pyqtSignal(str)

    def __init__(self, store: MonitorManager, parent=None):
        super().__init__(parent)
        self.store = store
        self.tasks: Dict[str, MonitorTask] = {}
        self.current_task_id: Optional[str] = None
        self._one_shot_worker: Optional[MonitorWorker] = None
        self._poll_worker: Optional[MonitorWorker] = None
        self._log_max_applied = False
        self._loading_form = False
        # 接收 MainWindow 的统一日志回调，外部赋值（复用定时页主面板汇总）
        self.log_callback = None
        self._worker_log_signal.connect(self._on_worker_log)
        self.filter_column_loaded = False
        self._col_combos: Dict[int, QComboBox] = {}
        self.init_ui()
        self.connect_signals()
        self.reload_tasks()
        # 每次启动都重建常驻轮询线程，保证与 UI 同生命周期的干净状态
        self._start_polling()

    # ----------------------------- UI -----------------------------
    def init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(8)

        # --- 左侧：任务列表 + 操作按钮 ---
        left_panel = QWidget()
        left_column = QVBoxLayout(left_panel)
        left_column.setContentsMargins(0, 0, 0, 0)
        left_column.setSpacing(6)
        left_bar = ChipBar(exclusive=False)
        left_column.addWidget(left_bar)

        list_group = left_bar.add_section("📋 监控配置", collapsed=True)
        left_layout = list_group.contentLayout()

        self.task_list = QListWidget()
        self.task_list.setMinimumWidth(240)
        left_layout.addWidget(self.task_list)

        left_btn1 = QHBoxLayout()
        self.add_task_btn = QPushButton("新建配置")
        self.import_btn = QPushButton("导入配置")
        left_btn1.addWidget(self.add_task_btn)
        left_btn1.addWidget(self.import_btn)
        left_layout.addLayout(left_btn1)

        left_btn2 = QHBoxLayout()
        self.delete_btn = QPushButton("删除配置")
        self.toggle_btn = QPushButton("启用/禁用")
        left_btn2.addWidget(self.delete_btn)
        left_btn2.addWidget(self.toggle_btn)
        left_layout.addLayout(left_btn2)

        self.run_once_btn = QPushButton("立即执行一次")
        left_layout.addWidget(self.run_once_btn)

        self.poll_status_label = QLabel("轮询状态: 运行中")
        self.poll_status_label.setStyleSheet("color:#666; font-size:11px;")
        left_layout.addWidget(self.poll_status_label)

        # 关键：把「监控配置」ChipSection 加入左侧列布局，否则面板是孤儿不可见
        left_column.addWidget(list_group)
        left_column.addStretch()
        main_layout.addWidget(left_panel, 0)

        # --- 右侧：编辑表单（芯片排，可同开，简洁纵向布局） ---
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)
        right_bar = ChipBar(exclusive=True)  # 手风琴：右栏设置一次只展开一个
        right_layout.addWidget(right_bar)

        # 1. 基本设置
        base_group = right_bar.add_section("📌 基本设置", collapsed=True)
        base_layout = base_group.contentLayout()
        row = QHBoxLayout()
        row.addWidget(QLabel("名称:"))
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("例如：群文件监控")
        row.addWidget(self.name_edit, 1)
        base_layout.addLayout(row)

        row = QHBoxLayout()
        row.addWidget(QLabel("监控目录:"))
        self.watch_path_edit = QLineEdit()
        row.addWidget(self.watch_path_edit, 1)
        self.browse_btn = QPushButton("浏览")
        row.addWidget(self.browse_btn)
        base_layout.addLayout(row)

        row = QHBoxLayout()
        row.addWidget(QLabel("文件匹配:"))
        self.pattern_edit = QLineEdit("*.xlsx;*.xls")
        self.pattern_edit.setPlaceholderText("多个用分号;分隔，如 *.xlsx;*.xls")
        row.addWidget(self.pattern_edit, 1)
        base_layout.addLayout(row)

        row = QHBoxLayout()
        row.addWidget(QLabel("文件前缀:"))
        self.file_prefix_edit = QLineEdit()
        self.file_prefix_edit.setPlaceholderText("留空匹配全部文件；填如 主干及分支 只处理同名开头的文件")
        row.addWidget(self.file_prefix_edit, 1)
        base_layout.addLayout(row)

        row = QHBoxLayout()
        row.addWidget(QLabel("扫描周期(分钟):"))
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 1440)
        self.interval_spin.setValue(5)
        row.addWidget(self.interval_spin, 1)
        base_layout.addLayout(row)

        self.include_subdir_check = QCheckBox("包含子文件夹中的匹配文件")
        self.include_subdir_check.setChecked(False)
        self.include_subdir_check.setToolTip("勾选后递归扫描监控目录的子文件夹（默认仅当前层）")
        base_layout.addWidget(self.include_subdir_check)
        right_layout.addWidget(base_group)

        # 2. 筛选与对比
        filter_group = right_bar.add_section("🔍 筛选与对比", collapsed=True)
        filter_layout = filter_group.contentLayout()

        row = QHBoxLayout()
        row.addWidget(QLabel("Sheet名称:"))
        self.sheet_edit = QLineEdit()
        self.sheet_edit.setPlaceholderText("留空 = 第一个 sheet")
        row.addWidget(self.sheet_edit, 1)
        self.read_header_btn = QPushButton("读取表头")
        row.addWidget(self.read_header_btn)
        filter_layout.addLayout(row)

        row = QHBoxLayout()
        row.addWidget(QLabel("筛选列:"))
        self.filter_column_combo = QComboBox()
        self.filter_column_combo.setEditable(True)
        self.filter_column_combo.setPlaceholderText("留空 = 全表")
        row.addWidget(self.filter_column_combo, 1)
        filter_layout.addLayout(row)

        filter_layout.addWidget(QLabel("筛选值(每行一个 / 逗号分隔):"))
        self.filter_values_edit = QTextEdit()
        self.filter_values_edit.setMaximumHeight(80)
        self.filter_values_edit.setPlaceholderText("例如：\n北京\n上海\n或：北京,上海")
        filter_layout.addWidget(self.filter_values_edit)

        self.compare_check = QCheckBox("开启内容对比(增量，只推送新增)")
        self.compare_check.setChecked(True)
        self.no_compare_check = QCheckBox("关闭对比(有新文件即推送)")
        self.compare_check.toggled.connect(self._on_compare_toggled)
        self.no_compare_check.toggled.connect(self._on_compare_toggled)
        filter_layout.addWidget(self.compare_check)
        filter_layout.addWidget(self.no_compare_check)
        right_layout.addWidget(filter_group)

        # 3. 推送内容
        send_group = right_bar.add_section("📤 推送内容", collapsed=True)
        send_layout = send_group.contentLayout()
        title_row = QHBoxLayout()
        title_row.addWidget(QLabel("文字标题:"))
        self.text_title_edit = QLineEdit()
        self.text_title_edit.setPlaceholderText("留空 = 默认「监控新增提醒」")
        title_row.addWidget(self.text_title_edit, 1)
        send_layout.addLayout(title_row)
        self.send_text_check = QCheckBox("推送文字说明")
        self.send_text_check.setChecked(True)
        self.text_detail_check = QCheckBox("文字附带逐行明细")
        self.text_detail_check.setChecked(True)
        self.send_image_check = QCheckBox("推送图片")
        self.send_file_check = QCheckBox("推送表格文件")
        self.monitor_minimize_check = QCheckBox("发送后最小化微信")
        self.monitor_minimize_check.setChecked(True)
        self.monitor_minimize_check.setToolTip("监控推送完成后最小化微信窗口（隐私保护）；取消勾选则保持微信窗口不变")
        send_layout.addWidget(self.send_text_check)
        send_layout.addWidget(self.text_detail_check)
        send_layout.addWidget(self.send_image_check)
        send_layout.addWidget(self.send_file_check)
        send_layout.addWidget(self.monitor_minimize_check)

        # 合并发送顺序（文字/图片/文件）：决定两步粘贴先后与文件卡片排列（未勾选内容自动跳过）
        order_row = QHBoxLayout()
        order_row.addWidget(QLabel("发送顺序:"))
        self.send_order_combo = QComboBox()
        self.send_order_combo.addItem("文字 → 图片 → 文件", ["text", "image", "attachment"])
        self.send_order_combo.addItem("文字 → 文件 → 图片", ["text", "attachment", "image"])
        self.send_order_combo.addItem("图片 → 文字 → 文件", ["image", "text", "attachment"])
        self.send_order_combo.addItem("图片 → 文件 → 文字", ["image", "attachment", "text"])
        self.send_order_combo.addItem("文件 → 文字 → 图片", ["attachment", "text", "image"])
        self.send_order_combo.addItem("文件 → 图片 → 文字", ["attachment", "image", "text"])
        self.send_order_combo.setToolTip("合并发送时消息里的排列顺序：文字在前则先粘文字再粘文件卡片；文件间按图片/文件先后一次粘贴")
        order_row.addWidget(self.send_order_combo, 1)
        send_layout.addLayout(order_row)

        # 写入智能表格（AirScript webhook，独立可选出口）：勾选后显示配置。
        self.airsync_check = QCheckBox("写入智能表格(AirScript)")
        self.airsync_check.setToolTip(
            "把清洗/提取后的数据（去掉首行标题）追加写入自己拥有的金山智能表格\n"
            "指定 sheet 的指定列，从最后一个非空行往下逐行写入。\n"
            "凭证使用脚本令牌 AirScript-Token（webhook 请求头），与现有 wps_sid 读取链路完全独立。")
        send_layout.addWidget(self.airsync_check)

        self.airsync_detail = QWidget()
        adl = QVBoxLayout(self.airsync_detail)
        adl.setContentsMargins(0, 0, 0, 0)

        row_web = QHBoxLayout()
        row_web.addWidget(QLabel("Webhook:"))
        self.airsync_webhook_edit = QLineEdit()
        self.airsync_webhook_edit.setPlaceholderText("脚本 webhook 链接（脚本编辑器复制）")
        row_web.addWidget(self.airsync_webhook_edit, 1)
        adl.addLayout(row_web)

        row_tok = QHBoxLayout()
        row_tok.addWidget(QLabel("脚本令牌:"))
        self.airsync_token_edit = QLineEdit()
        self.airsync_token_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.airsync_token_edit.setPlaceholderText("AirScript-Token（脚本编辑器盾牌图标创建）")
        row_tok.addWidget(self.airsync_token_edit, 1)
        adl.addLayout(row_tok)

        row_sheet = QHBoxLayout()
        row_sheet.addWidget(QLabel("目标Sheet:"))
        self.airsync_sheet_edit = QLineEdit()
        self.airsync_sheet_edit.setPlaceholderText("留空 = 表内活动表")
        row_sheet.addWidget(self.airsync_sheet_edit, 1)
        adl.addLayout(row_sheet)

        row_col = QHBoxLayout()
        row_col.addWidget(QLabel("起始列:"))
        self.airsync_start_col_spin = QSpinBox()
        self.airsync_start_col_spin.setRange(1, 702)
        self.airsync_start_col_spin.setValue(1)
        self.airsync_start_col_spin.setToolTip("数据追加写入的起始列，1=A，2=B …")
        row_col.addWidget(self.airsync_start_col_spin)
        row_col.addSpacing(8)
        row_col.addWidget(QLabel("列数:"))
        self.airsync_col_count_spin = QSpinBox()
        self.airsync_col_count_spin.setRange(0, 200)
        self.airsync_col_count_spin.setValue(0)
        self.airsync_col_count_spin.setSpecialValueText("全部")
        self.airsync_col_count_spin.setToolTip("0=写入每行全部列；>0=只写前 N 列")
        row_col.addWidget(self.airsync_col_count_spin)
        row_col.addStretch(1)
        adl.addLayout(row_col)

        self.airsync_hint = QLabel("写入列 = 清洗后提取列；去首行标题，从末尾非空行往下追加。")
        self.airsync_hint.setWordWrap(True)
        self.airsync_hint.setStyleSheet("color: #808080;")
        adl.addWidget(self.airsync_hint)

        self.airsync_detail.setVisible(False)
        send_layout.addWidget(self.airsync_detail)
        self.airsync_check.toggled.connect(self._on_airsync_toggled)

        # 清洗后条数判断拦截：len(clean_rows) 与 阈值 满足 操作符 关系时，本次不发送
        self.limit_enabled_check = QCheckBox("条数判断拦截")
        self.limit_enabled_check.setToolTip("对清洗后待发送的行数做判断：满足【大于/小于/等于】设定条数时本次不发送。基线照常推进")
        send_layout.addWidget(self.limit_enabled_check)
        limit_row = QHBoxLayout()
        self.limit_op_combo = QComboBox()
        self.limit_op_combo.addItem("条数不满足时不拦截(关闭)", "")
        self.limit_op_combo.addItem("条数大于阈值 → 不发送", ">")
        self.limit_op_combo.addItem("条数小于阈值 → 不发送", "<")
        self.limit_op_combo.addItem("条数等于阈值 → 不发送", "==")
        self.limit_op_combo.setEnabled(False)
        limit_row.addWidget(QLabel("清洗后"))
        limit_row.addWidget(self.limit_op_combo, 1)
        self.limit_count_spin = QSpinBox()
        self.limit_count_spin.setRange(0, 999999)
        self.limit_count_spin.setValue(0)
        self.limit_count_spin.setEnabled(False)
        limit_row.addWidget(self.limit_count_spin)
        limit_row.addWidget(QLabel("条"))
        self.limit_enabled_check.toggled.connect(
            lambda on: (self.limit_op_combo.setEnabled(on),
                        self.limit_count_spin.setEnabled(on)))
        send_layout.addLayout(limit_row)

        # 指定区域截图（与筛选/清洗/提取列并存、独立勾选；仅「关闭对比」模式生效）
        self.snapshot_check = QCheckBox("指定区域截图（仅关闭对比时生效）")
        self.snapshot_check.setToolTip("勾选后，图片内容改为截取 Sheet 指定区域（留空范围 = 整表已用区域）")
        self.snapshot_check.setChecked(False)
        row = QHBoxLayout()
        row.addWidget(QLabel("截图区域:"))
        self.snapshot_range_edit = QLineEdit()
        self.snapshot_range_edit.setPlaceholderText("如 A1:F20；留空 = 整表已用区域")
        row.addWidget(self.snapshot_range_edit, 1)
        self.snapshot_hint = QLabel("⚠ 指定区域截图仅在「关闭对比」模式下生效")
        self.snapshot_hint.setStyleSheet("color:#e08a00; font-size:11px;")
        self.snapshot_hint.setVisible(False)
        send_layout.addWidget(self.snapshot_check)
        send_layout.addLayout(row)
        send_layout.addWidget(self.snapshot_hint)
        right_layout.addWidget(send_group)

        # 3.5. 列设置：清洗列 + 发送保留列 + 对比列（统一面板）
        col_group = right_bar.add_section("🧰 列设置", collapsed=True)
        col_layout = col_group.contentLayout()

        col_layout.addWidget(QLabel("清洗列（要清洗的列，可多选/手动添加）:"))
        self.clean_cols_list = QListWidget()
        self.clean_cols_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.clean_cols_list.setMaximumHeight(90)
        col_layout.addWidget(self.clean_cols_list)
        self._add_col_input_row(col_layout, self.clean_cols_list)

        row = QHBoxLayout()
        row.addWidget(QLabel("清洗方式:"))
        self.strip_space_check = QCheckBox("去空格")
        self.strip_space_check.setChecked(True)
        self.drop_empty_check = QCheckBox("去空(空剔除)")
        self.drop_empty_check.setToolTip("勾选后，该列清洗后为空的单元格所在行按下方规则剔除/保留")
        self.digits_check = QCheckBox("去非数字")
        row.addWidget(self.strip_space_check)
        row.addWidget(self.drop_empty_check)
        row.addWidget(self.digits_check)
        col_layout.addLayout(row)

        row = QHBoxLayout()
        row.addWidget(QLabel("取前N位:"))
        self.take_first_spin = QSpinBox()
        self.take_first_spin.setRange(0, 50)
        self.take_first_spin.setValue(0)
        self.take_first_spin.setToolTip("0 = 不截取")
        row.addWidget(self.take_first_spin)
        row.addWidget(QLabel("前缀(须相符):"))
        self.prefix_edit = QLineEdit()
        self.prefix_edit.setPlaceholderText("空 = 不校验前缀")
        row.addWidget(self.prefix_edit, 1)
        col_layout.addLayout(row)

        row = QHBoxLayout()
        row.addWidget(QLabel("前缀不符时:"))
        self.clean_fail_combo = QComboBox()
        self.clean_fail_combo.addItem("剔除该行", "drop")
        self.clean_fail_combo.addItem("保留", "keep")
        row.addWidget(self.clean_fail_combo, 1)
        col_layout.addLayout(row)

        col_layout.addWidget(QLabel("发送保留列（提取/只写这些列；留空 = 全部列）:"))
        self.extract_cols_list = QListWidget()
        self.extract_cols_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.extract_cols_list.setMaximumHeight(90)
        col_layout.addWidget(self.extract_cols_list)
        self._add_col_input_row(col_layout, self.extract_cols_list)

        col_layout.addWidget(QLabel("对比列（决定「新增」依据；留空 = 比对全部列）:"))
        self.compare_cols_list = QListWidget()
        self.compare_cols_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.compare_cols_list.setMaximumHeight(90)
        col_layout.addWidget(self.compare_cols_list)
        self._add_col_input_row(col_layout, self.compare_cols_list)

        right_layout.addWidget(col_group)

        # 4. 微信接收人
        recv_group = right_bar.add_section("👤 微信接收人", collapsed=True)
        recv_layout = recv_group.contentLayout()
        self.recipients_edit = QTextEdit()
        self.recipients_edit.setMaximumHeight(90)
        self.recipients_edit.setPlaceholderText("每个微信昵称一行，或用逗号分隔")
        recv_layout.addWidget(self.recipients_edit)
        right_layout.addWidget(recv_group)

        # 5. 时间窗
        time_group = right_bar.add_section("🕐 时间窗", collapsed=True)
        time_layout = time_group.contentLayout()
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("重复模式:"))
        self.repeat_mode_combo = QComboBox()
        self.repeat_mode_combo.addItem("每天", "daily")
        self.repeat_mode_combo.addItem("每周指定日期", "weekly")
        self.repeat_mode_combo.addItem("指定日期(一次)", "once")
        row1.addWidget(self.repeat_mode_combo, 1)
        time_layout.addLayout(row1)

        self.weekday_group_box = QGroupBox("选择周几(每周模式生效):")
        weekday_layout = QHBoxLayout(self.weekday_group_box)
        self.weekday_checks: Dict[int, QCheckBox] = {}
        for idx, name in enumerate(WEEKDAY_NAMES, start=1):
            cb = QCheckBox(name)
            self.weekday_checks[idx] = cb
            weekday_layout.addWidget(cb)
        time_layout.addWidget(self.weekday_group_box)

        self.date_group_box = QGroupBox("执行日期(一次性模式生效，YYYY-MM-DD 逗号分隔):")
        date_layout = QVBoxLayout(self.date_group_box)
        self.run_dates_edit = QLineEdit()
        self.run_dates_edit.setPlaceholderText("例如：2026-10-01,2026-10-02")
        date_layout.addWidget(self.run_dates_edit)
        time_layout.addWidget(self.date_group_box)
        self._sync_time_window(self.repeat_mode_combo.currentIndex())

        # 每天运行时间段（限定在起止时间内才监控/推送；起止相同或未勾选=全天）
        time_row = QHBoxLayout()
        self.active_window_check = QCheckBox("限定每天时段")
        self.active_window_check.setToolTip("勾选后，只在起止时间段内扫描并推送；\n未勾选 = 全天监控")
        time_row.addWidget(self.active_window_check)
        time_row.addWidget(QLabel("起:"))
        self.active_start_edit = QTimeEdit()
        self.active_start_edit.setDisplayFormat("HH:mm")
        self.active_start_edit.setTime(QTime(7, 0))
        time_row.addWidget(self.active_start_edit)
        time_row.addWidget(QLabel("止:"))
        self.active_end_edit = QTimeEdit()
        self.active_end_edit.setDisplayFormat("HH:mm")
        self.active_end_edit.setTime(QTime(22, 0))
        self.active_end_edit.setToolTip("起止相同 = 全天")
        time_row.addWidget(self.active_end_edit)
        self.active_start_edit.setEnabled(False)
        self.active_end_edit.setEnabled(False)
        self.active_window_check.toggled.connect(self.active_start_edit.setEnabled)
        self.active_window_check.toggled.connect(self.active_end_edit.setEnabled)
        time_layout.addLayout(time_row)

        right_layout.addWidget(time_group)

        # 6. 保存
        save_group = right_bar.add_section("💾 保存", collapsed=True)
        save_layout = save_group.contentLayout()
        save_row = QHBoxLayout()
        self.save_btn = QPushButton("保存配置")
        self.clear_btn = QPushButton("清空表单")
        save_row.addWidget(self.save_btn)
        save_row.addWidget(self.clear_btn)
        save_layout.addLayout(save_row)
        right_layout.addWidget(save_group)

        # 运行日志 + 停止轮询
        log_group = right_bar.add_section("📜 运行日志")
        log_layout = log_group.contentLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(120)
        self.log_text.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_text)
        log_btn_row = QHBoxLayout()
        self.stop_poll_btn = QPushButton("停止轮询")
        log_btn_row.addWidget(self.stop_poll_btn)
        log_btn_row.addStretch()
        log_layout.addLayout(log_btn_row)
        right_layout.addWidget(log_group)

        right_layout.addStretch()
        main_layout.addWidget(right, 1)

        # 注册所有可折叠分组，供全局精简/详细模式遍历（含左侧配置列表，与定时发送一致）
        self._collapsible_groups = [
            list_group, base_group, filter_group, send_group,
            col_group, recv_group, time_group, save_group, log_group,
        ]
        self.list_group = list_group
        self.save_group = save_group
        # 表单脏检测：登录后 _form_ready=True 才生效（防止程序化填充误判为改动）
        self._form_dirty = False
        self._form_ready = False
        # 监控发送失败常驻通知去重：记录当前处于"失败未恢复"状态的任务名，
        # 仅在该任务由成功转入失败时弹一次，避免轮询反复失败导致通知堆积
        self._monitor_fail_notified: set = set()

    def connect_signals(self):
        self.task_list.currentItemChanged.connect(self._on_task_selected)
        self.add_task_btn.clicked.connect(self._on_add_task)
        self.delete_btn.clicked.connect(self._on_delete_task)
        self.toggle_btn.clicked.connect(self._on_toggle_enabled)
        self.run_once_btn.clicked.connect(self._on_run_once)
        self.browse_btn.clicked.connect(self._on_browse)
        self.read_header_btn.clicked.connect(self._on_read_headers)
        self.repeat_mode_combo.currentIndexChanged.connect(self._sync_time_window)
        self.save_btn.clicked.connect(self._on_save)
        self.clear_btn.clicked.connect(self._reset_form)
        self.stop_poll_btn.clicked.connect(self._on_toggle_polling)
        self.import_btn.clicked.connect(self._on_import_profile)
        self.snapshot_check.toggled.connect(self._on_snapshot_toggled)
        self._connect_form_dirty_signals()
        self._form_ready = True

    # ------------------------- 表单脏检测（改动自动弹开保存栏） -------------------------
    def _connect_form_dirty_signals(self):
        """所有表单控件改动 → 置脏并自动展开「💾 保存」栏（不挤掉正在编辑的栏）。"""
        mark = self._mark_form_dirty
        self.name_edit.textChanged.connect(mark)
        self.watch_path_edit.textChanged.connect(mark)
        self.pattern_edit.textChanged.connect(mark)
        self.file_prefix_edit.textChanged.connect(mark)
        self.interval_spin.valueChanged.connect(mark)
        self.sheet_edit.textChanged.connect(mark)
        self.filter_column_combo.currentTextChanged.connect(mark)
        self.filter_values_edit.textChanged.connect(mark)
        self.compare_check.stateChanged.connect(mark)
        self.no_compare_check.stateChanged.connect(mark)
        self.send_text_check.stateChanged.connect(mark)
        self.send_image_check.stateChanged.connect(mark)
        self.send_file_check.stateChanged.connect(mark)
        self.monitor_minimize_check.stateChanged.connect(mark)
        self.limit_enabled_check.stateChanged.connect(mark)
        self.limit_op_combo.currentIndexChanged.connect(mark)
        self.limit_count_spin.valueChanged.connect(mark)
        self.send_order_combo.currentIndexChanged.connect(mark)
        self.snapshot_check.stateChanged.connect(mark)
        self.snapshot_range_edit.textChanged.connect(mark)
        self.airsync_check.stateChanged.connect(mark)
        self.airsync_webhook_edit.textChanged.connect(mark)
        self.airsync_token_edit.textChanged.connect(mark)
        self.airsync_sheet_edit.textChanged.connect(mark)
        self.airsync_start_col_spin.valueChanged.connect(mark)
        self.airsync_col_count_spin.valueChanged.connect(mark)
        self.text_title_edit.textChanged.connect(mark)
        self.text_detail_check.stateChanged.connect(mark)
        self.include_subdir_check.stateChanged.connect(mark)
        self.strip_space_check.stateChanged.connect(mark)
        self.drop_empty_check.stateChanged.connect(mark)
        self.digits_check.stateChanged.connect(mark)
        self.take_first_spin.valueChanged.connect(mark)
        self.prefix_edit.textChanged.connect(mark)
        self.clean_fail_combo.currentIndexChanged.connect(mark)
        for lw in (self.clean_cols_list, self.extract_cols_list,
                   self.compare_cols_list):
            lw.itemChanged.connect(lambda *a: mark())
        self.recipients_edit.textChanged.connect(mark)
        self.repeat_mode_combo.currentIndexChanged.connect(mark)
        for cb in self.weekday_checks.values():
            cb.stateChanged.connect(mark)
        self.run_dates_edit.textChanged.connect(mark)
        self.active_window_check.stateChanged.connect(mark)
        self.active_start_edit.timeChanged.connect(mark)
        self.active_end_edit.timeChanged.connect(mark)

    def _mark_form_dirty(self, *_args):
        """表单项被用户改动：置脏并自动展开「保存」栏（程序填充阶段忽略）。"""
        if not self._form_ready or self._loading_form:
            return
        self._form_dirty = True
        if self.save_group.is_collapsed():
            self.save_group.set_collapsed(False, animate=True, accordion_close=False)

    def _clear_form_dirty(self):
        """保存/重置成功后清脏并可选收起保存栏。"""
        self._form_dirty = False

    # ----------------------------- 任务列表 -----------------------------
    def reload_tasks(self):
        self.task_list.blockSignals(True)
        self.task_list.clear()
        tasks = self.store.load_all()
        self.tasks = {t.id: t for t in tasks}
        for t in tasks:
            item = QListWidgetItem(self._format_visual_label(t))
            item.setData(Qt.ItemDataRole.UserRole, t.id)
            self.task_list.addItem(item)
        self.task_list.blockSignals(False)
        if tasks:
            self.task_list.setCurrentRow(0)

    def _format_visual_label(self, task: MonitorTask) -> str:
        mark = "✅" if task.enabled else "⬜"
        rule = {
            "weekly": "每周",
            "once": "一次",
        }.get(task.repeat_mode, "每天")
        target = f"人数:{len(task.recipients)}"
        return f"{mark} {task.name}  | {rule} | {target}"

    def _current_task_id_from_list(self) -> Optional[str]:
        item = self.task_list.currentItem()
        if item is None:
            return None
        return item.data(Qt.ItemDataRole.UserRole)

    def _on_task_selected(self, current, previous):
        if not self._loading_form:
            self.current_task_id = (current.data(Qt.ItemDataRole.UserRole)
                                    if current else None)
        tid = self.current_task_id
        if tid and tid in self.tasks:
            self._load_task_to_form(self.tasks[tid])

    # ----------------------------- 时间窗 -----------------------------
    def _sync_time_window(self, index):
        mode = self.repeat_mode_combo.itemData(index)
        self.weekday_group_box.setVisible(mode == "weekly")
        self.date_group_box.setVisible(mode == "once")

    # ----------------------------- 对比互斥 -----------------------------
    def _on_compare_toggled(self, _checked):
        # 手动互斥：勾选一个时清掉另一个的勾选
        if self.compare_check.isChecked() and self.no_compare_check.isChecked():
            if self.sender() is self.compare_check:
                self.no_compare_check.setChecked(False)
            else:
                self.compare_check.setChecked(False)
        self._refresh_snapshot_hint()

    # ----------------------------- 指定区域截图 -----------------------------
    def _on_snapshot_toggled(self, checked: bool):
        if self._loading_form:
            return
        # 勾选指定区域截图时自动开启「推送图片」，保证截图能被发出去（取消勾选不影响图片开关）
        if checked and not self.send_image_check.isChecked():
            self.send_image_check.setChecked(True)
        self._refresh_snapshot_hint()

    def _on_airsync_toggled(self, _checked=None):
        # 勾选「写入智能表格」才显示其配置区（程序填充阶段无需额外处理）
        self.airsync_detail.setVisible(self.airsync_check.isChecked())

    def _refresh_snapshot_hint(self):
        # 开启对比时提示该功能不生效（与关闭对比模式并存，用户自由选择）
        self.snapshot_hint.setVisible(
            self.compare_check.isChecked() and self.snapshot_check.isChecked()
        )

    # ----------------------------- 浏览目录 / 读取表头 -----------------------------
    def _on_browse(self):
        path = QFileDialog.getExistingDirectory(
            self, "选择监控目录", self.watch_path_edit.text().strip())
        if path:
            self.watch_path_edit.setText(path)

    def _on_read_headers(self):
        watch_path = self.watch_path_edit.text().strip()
        if not watch_path or not os.path.isdir(watch_path):
            self.log("请先填写有效的监控目录")
            return
        patterns = [p.strip() for p in
                    (self.pattern_edit.text() or "*.*").split(";") if p.strip()]
        file_path = None
        for pat in patterns or ["*.*"]:
            hits = glob.glob(os.path.join(watch_path, pat))
            for h in hits:
                if os.path.isfile(h):
                    file_path = h
                    break
            if file_path:
                break
        if not file_path:
            self.log("目录下没有匹配的文件，无法读取表头")
            return
        try:
            from modules import monitor_engine
            headers, _rows = monitor_engine.read_sheet_rows(
                file_path, sheet_name=self.sheet_edit.text().strip())
            self.filter_column_combo.clear()
            self.filter_column_combo.addItem("")  # 空=全表
            for h in headers:
                self.filter_column_combo.addItem(h)
            # 同步填充三个列设置列表（保留当前选中）
            keep = {
                "clean": self._selected_cols(self.clean_cols_list),
                "extract": self._selected_cols(self.extract_cols_list),
                "compare": self._selected_cols(self.compare_cols_list),
            }
            self._fill_columns_list(self.clean_cols_list, headers, keep["clean"])
            self._fill_columns_list(self.extract_cols_list, headers, keep["extract"])
            self._fill_columns_list(self.compare_cols_list, headers, keep["compare"])
            # 表头回填到各下拉候选（去重保留已有项表达式）
            for lw in (self.clean_cols_list, self.extract_cols_list,
                       self.compare_cols_list):
                combo = self._col_combos.get(id(lw))
                if combo is None:
                    continue
                seen = set(combo.itemText(i) for i in range(combo.count()))
                for h in headers:
                    if str(h).strip() and str(h) not in seen:
                        combo.addItem(str(h))
                        seen.add(str(h))
            self.filter_column_loaded = True
            if self.sheet_edit.text().strip():
                self.log(f"已读取 {os.path.basename(file_path)} 的表头 {len(headers)} 列")
            else:
                self.log(f"已读取 {os.path.basename(file_path)} 首个 sheet 的表头 {len(headers)} 列")
        except Exception as exc:
            self.log(f"读取表头失败: {exc}")

    # ----------------------------- 列设置(清洗/保留/对比) helper -----------------------------
    def _add_col_input_row(self, column_layout, list_widget):
        """为某个列列表添加『输入列名 + 添加/移除』一行，支持手动填写列名。

        不依赖"读取表头"即可录入列名：在可编辑下拉里输入列名回车或点「＋」，
        会加入对应列表并自动选中；选中列表项后点「−」可移除。
        """
        row = QHBoxLayout()
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        combo.setPlaceholderText("输入或选择列名…")
        add_btn = QPushButton("＋添加")
        add_btn.setFixedWidth(58)
        del_btn = QPushButton("−移除选中")
        del_btn.setFixedWidth(70)

        def _add_col():
            text = combo.currentText().strip()
            if not text:
                return
            # 已存在则不重复添加
            for i in range(list_widget.count()):
                if list_widget.item(i).text() == text:
                    list_widget.item(i).setSelected(True)
                    combo.clearEditText()
                    return
            item = QListWidgetItem(text)
            list_widget.addItem(item)
            item.setSelected(True)
            combo.clearEditText()
            # addItem 不会触发 itemChanged，需显式置脏以弹开保存栏
            self._mark_form_dirty()

        def _del_col():
            removed = False
            for i in list(range(list_widget.count() - 1, -1, -1)):
                if list_widget.item(i).isSelected():
                    list_widget.takeItem(i)
                    removed = True
            # takeItem 不会触发 itemChanged，需显式置脏以弹开保存栏
            if removed:
                self._mark_form_dirty()

        combo.lineEdit().returnPressed.connect(_add_col)
        add_btn.clicked.connect(_add_col)
        del_btn.clicked.connect(_del_col)
        row.addWidget(combo, 1)
        row.addWidget(add_btn)
        row.addWidget(del_btn)
        column_layout.addLayout(row)
        # 记录每个列表对应的下拉，供读取表头后回填候选
        self._col_combos.setdefault(id(list_widget), combo)

    def _selected_cols(self, list_widget) -> List[str]:
        """返回某个列设置列表的全部项。

        自 v1.4.2 后列表体只放用户显式添加/配置的列，故保存一律取列表全部内容，
        不再依赖 isSelected() 选中态——避免三个列表内容相同导致选中态互相污染串台。
        """
        return [list_widget.item(i).text() for i in range(list_widget.count())]

    def _fill_columns_list(self, list_widget, headers, selected):
        """填充某个列设置列表。

        关键设计：列表体只显示用户显式添加/已配置的列（保序去重），
        不再把全部表头塞进列表体。表头仅作为各列表下拉候选供用户挑选添加
        （在 _add_col_input_row / _on_read_headers 里维护）。这样三个列表
        各自只含自己配置的列，天然隔离、不会串台，且顺序 = 用户录入顺序。
        headers 参数保留仅为兼容调用点，不再用于列表体填充。
        """
        seen = set()
        order = []
        for c in (selected or []):
            cs = str(c).strip()
            if cs and cs not in seen:
                seen.add(cs)
                order.append(cs)
        list_widget.blockSignals(True)
        list_widget.clear()
        for name in order:
            it = QListWidgetItem(name)
            list_widget.addItem(it)
            it.setSelected(True)
        list_widget.blockSignals(False)

    def _current_headers(self) -> List[str]:
        """当前表头候选：已加载到筛选列下拉的表头（不含跨列表合并）。"""
        headers = []
        for i in range(self.filter_column_combo.count()):
            h = self.filter_column_combo.itemText(i).strip()
            if h and h not in headers:
                headers.append(h)
        return headers

    def _build_clean_rules(self, cols) -> List:
        from modules.monitor_config import CleanRule

        rules = []
        for c in cols:
            rules.append(CleanRule(
                column=c,
                strip_space=self.strip_space_check.isChecked(),
                drop_empty=self.drop_empty_check.isChecked(),
                keep_digits_only=self.digits_check.isChecked(),
                take_first_n=self.take_first_spin.value(),
                require_prefix=self.prefix_edit.text().strip(),
                on_fail=self.clean_fail_combo.currentData() or "drop",
            ))
        return rules

    def _on_import_profile(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "导入监控配置", "", "监控配置 (*.json)")
        if not path:
            return
        try:
            import json as _json
            from modules.monitor_config import MonitorTask as _MT

            with open(path, "r", encoding="utf-8") as fp:
                raw = _json.load(fp)
            task = _MT.from_dict(raw)
            task.id = task.id or ""
            self.store.save(task)          # 落成 profiles 独立配置
            self.current_task_id = task.id
            self.reload_tasks()
            self._set_list_selection(task.id)
            self._load_task_to_form(task)
            self.log(f"[监控] 已导入配置「{task.name}」")
        except Exception as exc:
            QMessageBox.warning(self, "导入失败", str(exc))

    # ----------------------------- 表单 <-> 任务 -----------------------------
    def _load_task_to_form(self, task: MonitorTask):
        self._loading_form = True
        try:
            self.name_edit.setText(task.name)
            self.watch_path_edit.setText(task.watch_path)
            self.pattern_edit.setText(task.file_pattern)
            self.file_prefix_edit.setText(task.file_prefix)
            self.interval_spin.setValue(task.poll_interval_min)
            self.sheet_edit.setText(task.sheet_name)
            self.filter_column_combo.setCurrentText(task.filter_column)
            self.filter_values_edit.setPlainText("\n".join(task.filter_values))
            self.compare_check.setChecked(task.compare_enabled)
            self.no_compare_check.setChecked(not task.compare_enabled)
            self.send_text_check.setChecked(task.send_text)
            self.send_image_check.setChecked(task.send_image)
            self.send_file_check.setChecked(task.send_file)
            self.monitor_minimize_check.setChecked(getattr(task, "minimize_after", True))
            self.limit_enabled_check.setChecked(task.limit_enabled)
            self.limit_op_combo.setEnabled(task.limit_enabled)
            self.limit_count_spin.setEnabled(task.limit_enabled)
            idx_lim = self.limit_op_combo.findData(task.limit_op)
            self.limit_op_combo.setCurrentIndex(idx_lim if idx_lim >= 0 else 0)
            self.limit_count_spin.setValue(task.limit_count)
            self.snapshot_check.setChecked(task.snapshot_enabled)
            self.snapshot_range_edit.setText(task.snapshot_range)
            self.text_title_edit.setText(task.text_title)
            self.text_detail_check.setChecked(task.text_detail)
            idx_order = self.send_order_combo.findData(
                getattr(task, "send_order", None))
            self.send_order_combo.setCurrentIndex(idx_order if idx_order >= 0 else 0)
            self.airsync_check.setChecked(task.airsync_enabled)
            self.airsync_webhook_edit.setText(task.airsync_webhook)
            self.airsync_token_edit.setText(task.airsync_token)
            self.airsync_sheet_edit.setText(task.airsync_sheet)
            self.airsync_start_col_spin.setValue(task.airsync_start_col)
            self.airsync_col_count_spin.setValue(task.airsync_col_count)
            self.airsync_detail.setVisible(task.airsync_enabled)
            self.include_subdir_check.setChecked(task.include_subdir)
            # 列设置：勾选清洗/保留/对比列，并应用统一清洗方式
            self.strip_space_check.setChecked(True)
            self.drop_empty_check.setChecked(False)
            self.digits_check.setChecked(False)
            self.take_first_spin.setValue(0)
            self.prefix_edit.clear()
            self.clean_fail_combo.setCurrentIndex(0)
            if task.clean_rules:
                r0 = task.clean_rules[0]
                self.strip_space_check.setChecked(r0.strip_space)
                self.drop_empty_check.setChecked(getattr(r0, "drop_empty", False))
                self.digits_check.setChecked(r0.keep_digits_only)
                self.take_first_spin.setValue(r0.take_first_n)
                self.prefix_edit.setText(r0.require_prefix)
                self.clean_fail_combo.setCurrentIndex(
                    0 if getattr(r0, "on_fail", "drop") != "keep" else 1)
            clean_cols = [r.column for r in task.clean_rules if getattr(r, "column", "")]
            self._fill_columns_list(self.clean_cols_list, self._current_headers(),
                                    clean_cols)
            self._fill_columns_list(self.extract_cols_list, self._current_headers(),
                                    task.extract_columns)
            self._fill_columns_list(self.compare_cols_list, self._current_headers(),
                                    task.compare_columns)
            self.recipients_edit.setPlainText("\n".join(task.recipients))
            mode_index = {"daily": 0, "weekly": 1, "once": 2}.get(
                task.repeat_mode, 0)
            self.repeat_mode_combo.setCurrentIndex(mode_index)
            for d, cb in self.weekday_checks.items():
                cb.setChecked(d in set(task.days))
            self.run_dates_edit.setText(",".join(task.run_dates))
            has_win = bool(task.active_start and task.active_end)
            self.active_window_check.setChecked(has_win)
            self.active_start_edit.setEnabled(has_win)
            self.active_end_edit.setEnabled(has_win)
            if task.active_start:
                self.active_start_edit.setTime(
                    QTime.fromString(task.active_start, "HH:mm"))
            if task.active_end:
                self.active_end_edit.setTime(
                    QTime.fromString(task.active_end, "HH:mm"))
            self._sync_time_window(self.repeat_mode_combo.currentIndex())
            self._refresh_snapshot_hint()
        finally:
            self._loading_form = False

    def _form_to_task(self, new_id: bool) -> MonitorTask:
        name = self.name_edit.text().strip()
        if not name:
            raise ValueError("请填写监控名称")
        watch_path = self.watch_path_edit.text().strip()
        if not watch_path:
            raise ValueError("请选择监控目录")

        existing_id = None if new_id else self.current_task_id
        task = MonitorTask(id=existing_id or "", name=name)
        task.enabled = True
        task.watch_path = watch_path
        task.file_pattern = self.pattern_edit.text().strip() or "*.xlsx;*.xls"
        task.file_prefix = self.file_prefix_edit.text().strip()
        task.poll_interval_min = self.interval_spin.value()
        task.sheet_name = self.sheet_edit.text().strip()
        task.filter_column = self.filter_column_combo.currentText().strip()
        task.filter_values = self._parse_filter_values(self.filter_values_edit.toPlainText())
        task.compare_enabled = self.compare_check.isChecked()
        task.send_text = self.send_text_check.isChecked()
        task.send_image = self.send_image_check.isChecked()
        task.send_file = self.send_file_check.isChecked()
        task.minimize_after = self.monitor_minimize_check.isChecked()
        task.limit_enabled = self.limit_enabled_check.isChecked()
        task.limit_op = self.limit_op_combo.currentData() or ""
        task.limit_count = self.limit_count_spin.value()
        task.snapshot_enabled = self.snapshot_check.isChecked()
        task.snapshot_range = self.snapshot_range_edit.text().strip()
        task.text_title = self.text_title_edit.text().strip()
        task.text_detail = self.text_detail_check.isChecked()
        task.send_order = list(
            self.send_order_combo.currentData() or ["text", "image", "attachment"])
        task.airsync_enabled = self.airsync_check.isChecked()
        task.airsync_webhook = self.airsync_webhook_edit.text().strip()
        task.airsync_token = self.airsync_token_edit.text().strip()
        task.airsync_sheet = self.airsync_sheet_edit.text().strip()
        task.airsync_start_col = self.airsync_start_col_spin.value()
        task.airsync_col_count = self.airsync_col_count_spin.value()
        task.include_subdir = self.include_subdir_check.isChecked()
        task.compare_columns = self._selected_cols(self.compare_cols_list)
        task.extract_columns = self._selected_cols(self.extract_cols_list)
        task.clean_rules = self._build_clean_rules(
            self._selected_cols(self.clean_cols_list))
        task.recipients = _parse_recipients(self.recipients_edit.toPlainText())
        task.repeat_mode = self.repeat_mode_combo.currentData() or "daily"
        task.days = sorted(d for d, cb in self.weekday_checks.items() if cb.isChecked())
        task.run_dates = [d.strip() for d in
                          self.run_dates_edit.text().replace("，", ",").split(",") if d.strip()]
        if self.active_window_check.isChecked():
            task.active_start = self.active_start_edit.time().toString("HH:mm")
            task.active_end = self.active_end_edit.time().toString("HH:mm")
        else:
            task.active_start = ""
            task.active_end = ""

        # 编辑已有任务时，保留历史 baseline/seen_files，避免重复推送。
        # 必须从磁盘读最新值：worker 轮询推进基线后只落盘，GUI 内存 self.tasks
        # 是加载时的旧快照——若用内存旧值覆盖保存，会冲掉 worker 刚推进的新基线，
        # 下一轮轮询把已推送过的行当"新增"重复推送（实测 07:20/07:25 同批数据重复发 3 次）。
        if existing_id:
            disk = self.store.get(existing_id)  # load_all 从磁盘读最新配置
            if disk is not None:
                task.baseline = dict(disk.baseline)
                task.seen_files = dict(disk.seen_files)
                task.created_at = disk.created_at
                task.enabled = disk.enabled
            elif existing_id in self.tasks:
                old = self.tasks[existing_id]
                task.baseline = dict(old.baseline)
                task.seen_files = dict(old.seen_files)
                task.created_at = old.created_at
                task.enabled = old.enabled
        return task

    def _parse_filter_values(self, text: str) -> List[str]:
        values: List[str] = []
        raw = str(text).replace("，", ",").replace(";", ",").replace("；", ",")
        for line in raw.splitlines():
            for part in line.split(","):
                v = part.strip()
                if v and v not in values:
                    values.append(v)
        return values

    # ----------------------------- 操作 -----------------------------
    def _on_add_task(self):
        self.current_task_id = None
        self._set_list_selection(None)
        self._reset_form()

    def _set_list_selection(self, tid: Optional[str]):
        self.task_list.blockSignals(True)
        self.task_list.clearSelection()
        if tid is not None:
            for i in range(self.task_list.count()):
                if self.task_list.item(i).data(Qt.ItemDataRole.UserRole) == tid:
                    self.task_list.setCurrentRow(i)
                    break
        self.task_list.blockSignals(False)

    def _on_delete_task(self):
        tid = self._current_task_id_from_list()
        if not tid:
            QMessageBox.information(self, "提示", "请先选择一个监控任务")
            return
        task = self.tasks.get(tid)
        reply = QMessageBox.question(
            self, "确认删除", f"确定删除监控「{task.name}」吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.store.delete(tid)
            self.log(f"[监控] 已删除任务: {task.name}")
            self.current_task_id = None
            self.reload_tasks()

    def _on_toggle_enabled(self):
        tid = self._current_task_id_from_list()
        if not tid:
            QMessageBox.information(self, "提示", "请先选择一个监控任务")
            return
        task = self.tasks.get(tid)
        new_state = not task.enabled
        self.store.set_enabled(tid, new_state)
        task.enabled = new_state
        self.log(f"[监控] 任务「{task.name}」已{'启用' if new_state else '禁用'}")
        self.reload_tasks()

    def _on_save(self):
        try:
            task = self._form_to_task(new_id=self.current_task_id is None)
        except ValueError as exc:
            QMessageBox.warning(self, "信息不完整", str(exc))
            return
        saved = self.store.save(task)
        self.log(f"[监控] 已保存监控「{saved.name}」")
        self._clear_form_dirty()
        self.current_task_id = saved.id
        self.reload_tasks()

    def _reset_form(self):
        self._loading_form = True
        try:
            self.name_edit.clear()
            self.watch_path_edit.clear()
            self.pattern_edit.setText("*.xlsx;*.xls")
            self.prefix_edit.clear()
            self.interval_spin.setValue(5)
            self.sheet_edit.clear()
            self.filter_column_combo.clear()
            self.filter_column_loaded = False
            self.filter_column_combo.setCurrentText("")
            self.filter_values_edit.clear()
            self.compare_check.setChecked(True)
            self.no_compare_check.setChecked(False)
            self.send_text_check.setChecked(True)
            self.send_image_check.setChecked(False)
            self.send_file_check.setChecked(False)
            self.monitor_minimize_check.setChecked(True)
            self.limit_enabled_check.setChecked(False)
            self.limit_op_combo.setCurrentIndex(0)
            self.limit_op_combo.setEnabled(False)
            self.limit_count_spin.setValue(0)
            self.limit_count_spin.setEnabled(False)
            self.snapshot_check.setChecked(False)
            self.snapshot_range_edit.clear()
            self.text_title_edit.clear()
            self.text_detail_check.setChecked(True)
            self.send_order_combo.setCurrentIndex(0)
            self.airsync_check.setChecked(False)
            self.airsync_webhook_edit.clear()
            self.airsync_token_edit.clear()
            self.airsync_sheet_edit.clear()
            self.airsync_start_col_spin.setValue(1)
            self.airsync_col_count_spin.setValue(0)
            self.airsync_detail.setVisible(False)
            self.include_subdir_check.setChecked(False)
            self.strip_space_check.setChecked(True)
            self.digits_check.setChecked(False)
            self.take_first_spin.setValue(0)
            self.prefix_edit.clear()
            self.clean_fail_combo.setCurrentIndex(0)
            for lw in (self.clean_cols_list, self.extract_cols_list,
                       self.compare_cols_list,):
                lw.blockSignals(True)
                lw.clear()
                lw.blockSignals(False)
            self.recipients_edit.clear()
            self.repeat_mode_combo.setCurrentIndex(0)
            for cb in self.weekday_checks.values():
                cb.setChecked(False)
            self.run_dates_edit.clear()
            self.active_window_check.setChecked(False)
            self.active_start_edit.setTime(QTime(7, 0))
            self.active_end_edit.setTime(QTime(22, 0))
            self.active_start_edit.setEnabled(False)
            self.active_end_edit.setEnabled(False)
            self._sync_time_window(0)
            self._refresh_snapshot_hint()
        finally:
            self._loading_form = False
        self._clear_form_dirty()

    # ----------------------------- 轮询 / 立即执行 -----------------------------
    def _start_polling(self):
        if self._poll_worker is not None and self._poll_worker.isRunning():
            return
        worker = MonitorWorker(store=self.store, one_shot=False, poll_floor_sec=30.0)
        self._bind_worker_signals(worker)
        self._poll_worker = worker
        self.poll_status_label.setText("轮询状态: 运行中")
        self.stop_poll_btn.setText("停止轮询")
        worker.start()
        self.log("[监控] 轮询已启动 (后台常驻)")

    def _on_toggle_polling(self):
        if self._poll_worker is not None and self._poll_worker.isRunning():
            self._poll_worker.stop()
            self.poll_status_label.setText("轮询状态: 已停止")
            self.stop_poll_btn.setText("开始轮询")
            self.log("[监控] 轮询已停止")
        else:
            self._start_polling()

    def _on_run_once(self):
        # 防重叠：仅防止「手动执行」自身重复执行。与常驻轮询并行是允许的，
        # 由 MonitorWorker 的任务级互斥保证同一任务不会同时被两个 worker 处理
        if self._one_shot_worker is not None and self._one_shot_worker.isRunning():
            self.log("正在执行，请稍候")
            return
        # 保存当前表单到所选任务（保留用户临时修改）
        tid = self._current_task_id_from_list()
        if not tid:
            QMessageBox.information(self, "提示", "请先选择要执行的监控任务")
            return
        try:
            task = self._form_to_task(new_id=False)
        except ValueError as exc:
            QMessageBox.warning(self, "信息不完整", str(exc))
            return
        # 回写并保存，让 worker 读到最新配置
        self.store.save(task)
        self.tasks[tid] = self.store.get(tid)

        worker = MonitorWorker(store=self.store, one_shot=True, poll_floor_sec=30.0)
        self._bind_worker_signals(worker)
        self._one_shot_worker = worker
        self.log(f"[监控] 手动执行一次: {task.name}")
        worker.start()

    def _bind_worker_signals(self, worker: MonitorWorker):
        worker.log.connect(self._worker_log_signal.emit)
        worker.task_finished.connect(self._on_task_finished)
        worker.run_finished.connect(self._on_run_finished)

    def _on_task_finished(self, task_name: str, success: int, failed: int):
        self.log(f"[监控] 任务「{task_name}」完成: 成功 {success} / 失败 {failed}")
        # 发送失败 → 右下角常驻通知（同定时任务：红色图标、不自动隐藏、点击置前）。
        # 仅在该任务由成功/无失败转入失败时弹一次；恢复成功后重新布防，避免反复失败堆积。
        if failed > 0:
            if task_name not in self._monitor_fail_notified:
                self._monitor_fail_notified.add(task_name)
                try:
                    ToastNotification.show_toast(
                        "监控发送失败",
                        f"监控「{task_name}」发送失败 {failed} 条，成功 {success} 条（点击查看）",
                        success=False, parent=self.window())
                except Exception:
                    logger.exception("[监控] 失败通知弹出失败")
        else:
            self._monitor_fail_notified.discard(task_name)

    def _on_run_finished(self):
        self.log("[监控] 本轮执行完毕")

    # ----------------------------- 日志 -----------------------------
    def log(self, message: str) -> None:
        if hasattr(self, "log_text") and self.log_text is not None:
            from datetime import datetime
            ts = datetime.now().strftime("%H:%M:%S")
            self.log_text.append(f"[{ts}] {message}")
            if not self._log_max_applied:
                self.log_text.document().setMaximumBlockCount(1200)
                self._log_max_applied = True
        logger.info(message)
        # 同时推送给 MainWindow，用于主界面统一日志面板
        if self.log_callback:
            try:
                self.log_callback(message)
            except Exception:
                pass

    def _on_worker_log(self, message: str) -> None:
        self.log(message)

    # 停止所有 worker（主窗口退出时调用）
    def shutdown(self):
        for attr in ("_poll_worker", "_one_shot_worker"):
            worker = getattr(self, attr, None)
            if worker is not None:
                try:
                    worker.stop()
                except Exception:
                    pass

    def closeEvent(self, event):
        self.shutdown()
        super().closeEvent(event)


class GlobalSettingsTab(QWidget):
    """「⚙️ 全量设置」主页签：应用级全局设置（跨所有 profile）。

    手风琴栏组（exclusive）容纳：
    - 「🎯 拟人节流」：全局开关 + 档位 + 自定义写动作随机延迟区间，
      关闭可显著加快批量发送（第三方驱动内置拟人节奏默认很慢）。
    - 「🔁 全量默认映射」：作为映射的**最后兜底层**。优先级：
      profile 自身映射 > 最近配置恢复的映射 > 全量默认映射 > 按原名发送。
    """

    # 全局设置被修改后发出，MainWindow 可据此同步应用（如重新 apply_rhythm）
    settings_changed = pyqtSignal()

    # (显示文本, 档位 key)；addItem(text, key) 存数据时以 key 为准
    _PROFILE_TEXT = [("关闭", "off"), ("快速", "fast"), ("自然", "natural"), ("保守", "calm")]

    def __init__(self, config_manager):
        super().__init__()
        self.config_manager = config_manager
        self.settings = config_manager.load_global_settings()
        # UI 回填期间屏蔽 onChange 触发，避免写入未初始化控件值
        self._loading = True
        # 关联的映射表文件路径（内存态，随 _collect_settings 持久化）
        self._map_linked_file = ""
        self._change_timer = QTimer(self)
        self._change_timer.setSingleShot(True)
        self._change_timer.setInterval(400)
        self._change_timer.timeout.connect(self._flush_and_emit)

        main_layout = QVBoxLayout(self)
        bar = ChipBar(exclusive=True)
        main_layout.addWidget(bar)

        # ---------------- 拟人节流 ----------------
        rhythm_group = bar.add_section("🎯 拟人节流", collapsed=True)
        ## add_section 只登记 section，面板必须由调用方 addWidget 到同列布局，
        ## 否则只有芯片按钮、点击不会显示内容区（与其余页面装配一致）
        main_layout.addWidget(rhythm_group)
        rl = rhythm_group.contentLayout()

        self.rhythm_enable_check = QCheckBox("启用拟人节流（关闭可显著提升批量发送速度）")
        rl.addWidget(self.rhythm_enable_check)

        row = QHBoxLayout()
        row.addWidget(QLabel("节流档位："))
        self.rhythm_profile_combo = QComboBox()
        for text, key in self._PROFILE_TEXT:
            self.rhythm_profile_combo.addItem(text, key)
        row.addWidget(self.rhythm_profile_combo)
        row.addWidget(QLabel("写动作随机间隔(秒)："))
        self.rhythm_min_spin = QDoubleSpinBox()
        self.rhythm_min_spin.setRange(0.0, 30.0)
        self.rhythm_min_spin.setDecimals(1)
        self.rhythm_min_spin.setSingleStep(0.5)
        row.addWidget(self.rhythm_min_spin)
        row.addWidget(QLabel("~"))
        self.rhythm_max_spin = QDoubleSpinBox()
        self.rhythm_max_spin.setRange(0.0, 30.0)
        self.rhythm_max_spin.setDecimals(1)
        self.rhythm_max_spin.setSingleStep(0.5)
        row.addWidget(self.rhythm_max_spin)
        row.addStretch()
        rl.addLayout(row)

        tip = QLabel(
            "拟人节流模拟真人操作节奏以降低微信风控风险。若批量发送等待过久，\n"
            "可将档位设为「快速」或手动把间隔调小，甚至直接关闭本开关。"
        )
        tip.setWordWrap(True)
        tip.setStyleSheet("color:#888;font-size:11px;")
        rl.addWidget(tip)

        # ---------------- 全量默认映射 ----------------
        map_group = bar.add_section("🔁 全量默认映射", collapsed=True)
        main_layout.addWidget(map_group)
        ml = map_group.contentLayout()

        self.map_enable_check = QCheckBox("启用全量默认映射（作为映射兜底层）")
        ml.addWidget(self.map_enable_check)

        drow = QHBoxLayout()
        drow.addWidget(QLabel("未命中映射时的兜底接收人："))
        self.map_default_edit = QLineEdit()
        self.map_default_edit.setPlaceholderText("留空则不使用兜底")
        drow.addWidget(self.map_default_edit)
        ml.addLayout(drow)

        # 映射表文件操作：导入 / 打开 / 关联，与「数据发送」的联系人映射一致
        mrow = QHBoxLayout()
        self.map_link_btn = QPushButton("🔗 关联映射表")
        self.map_import_btn = QPushButton("📥 导入映射表")
        self.map_open_btn = QPushButton("📂 打开映射表")
        self.map_link_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.map_import_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.map_open_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        mrow.addWidget(self.map_link_btn)
        mrow.addWidget(self.map_import_btn)
        mrow.addWidget(self.map_open_btn)
        self.map_file_label = QLabel("未关联映射表文件")
        self.map_file_label.setStyleSheet("color:#888;font-size:11px;")
        self.map_file_label.setWordWrap(True)
        mrow.addWidget(self.map_file_label, 1)
        mrow.addStretch()
        ml.addLayout(mrow)

        ml.addWidget(QLabel("映射列表（每行一个，格式：来源1;来源2=接收人1;接收人2）："))
        self.map_edit = QPlainTextEdit()
        self.map_edit.setPlaceholderText('张三;李四=王五\n"孙中枢，孙中枢"=孙中枢')
        self.map_edit.setFixedHeight(140)
        ml.addWidget(self.map_edit)

        tip2 = QLabel(
            "等号左侧可写多个来源名（用 ; 分隔），每个都映射到右侧的接收人列表；\n"
            "逗号属于名字本身，不用作分隔；可用\"\"包裹含分号的完整名字。\n"
            "优先级：配置自身映射 > 最近配置恢复的映射 > 全量默认映射 > 按原名发送。"
        )
        tip2.setWordWrap(True)
        tip2.setStyleSheet("color:#888;font-size:11px;")
        ml.addWidget(tip2)

        # 立即生效的信号连接
        self.rhythm_enable_check.toggled.connect(self._on_immediate_change)
        self.rhythm_profile_combo.currentIndexChanged.connect(self._on_immediate_change)
        # 文本类控件（含拖拽/粘贴）用防抖合并，避免一次编辑多次落盘
        self.map_enable_check.toggled.connect(self._on_debounced_change)
        self.map_default_edit.editingFinished.connect(self._on_debounced_change)
        self.map_edit.textChanged.connect(self._on_debounced_change)
        self.map_link_btn.clicked.connect(self._on_link_mapping_file)
        self.map_import_btn.clicked.connect(self._on_import_mapping_file)
        self.map_open_btn.clicked.connect(self._on_open_mapping_file)

        self._load_to_ui()
        self._loading = False

    # ---------------- UI <-> 设置 -----------------
    def _load_to_ui(self):
        rh = self.settings.get("rhythm") or {}
        self.rhythm_enable_check.setChecked(bool(rh.get("enabled", False)))
        idx = self.rhythm_profile_combo.findData(str(rh.get("profile", "off")))
        if idx < 0:
            idx = 0
        self.rhythm_profile_combo.setCurrentIndex(idx)
        self.rhythm_min_spin.setValue(float(rh.get("min_delay", 0.0)))
        self.rhythm_max_spin.setValue(float(rh.get("max_delay", 0.0)))

        dm = self.settings.get("default_mapping") or {}
        self.map_enable_check.setChecked(bool(dm.get("enabled", False)))
        self.map_default_edit.setText(str(dm.get("default_recipient", "") or ""))
        self._map_linked_file = str(dm.get("mapping_file", "") or "").strip()
        self._update_map_file_label()
        lines = []
        for m in (dm.get("mappings") or []):
            src = str(m.get("source_value", "") or "").strip()
            recips = [str(r).strip() for r in (m.get("recipients") or []) if str(r).strip()]
            if src and recips:
                lines.append(
                    f"{self._fmt_mapping_name(src)}="
                    f"{';'.join(self._fmt_mapping_name(r) for r in recips)}"
                )
        self.map_edit.setPlainText("\n".join(lines))

    @staticmethod
    def _fmt_mapping_name(name):
        """名字含分隔符/等号时用引号包裹，保证往返解析不歧义。"""
        if any(c in name for c in (";", "；", "=")):
            return f'"{name}"'
        return name

    @staticmethod
    def _split_mapping_names(text):
        """按分隔符切分名字；分隔符为 ; ；（引号内的分隔符不生效）。

        逗号**不是**分隔符——名字本身可能含逗号（如"孙中枢，孙中枢"）。
        返回去引号后的干净名字列表（含空串，由调用方过滤）。
        """
        parts = []
        buf = []
        in_quote = False
        for ch in text:
            if ch == '"':
                in_quote = not in_quote
                buf.append(ch)
            elif ch in (";", "；") and not in_quote:
                parts.append("".join(buf).strip())
                buf = []
            else:
                buf.append(ch)
        parts.append("".join(buf).strip())
        out = []
        for p in parts:
            p = p.strip()
            if p.startswith('"') and p.endswith('"') and len(p) >= 2:
                p = p[1:-1].strip()
            out.append(p)
        return out

    # ---------------- 映射表文件：关联 / 导入 / 打开 ----------------
    @staticmethod
    def _plain_split_values(raw):
        """把映射表文件单元格拆成多个值：按 / ; ；、。， 分隔，去空白空串。

        与 field 版 _split_rm_values 一致（普通联系人也用这套分隔符），
        保证「导入」与「数据发送」的映射表文件格式互通。

        引号包裹的整串视为一个完整值：其内部的分隔符（含逗号）不生效，
        并去掉两端引号——如 `"孙中枢,孙中枢"` 应保留为单个 `孙中枢,孙中枢`，
        而不是被拆成 `孙中枢` 与 `孙中枢"`（与 _split_mapping_names 行为一致）。
        """
        if not raw:
            return []
        text = str(raw).strip()
        if not text:
            return []
        seps = ("/", ";", "；", "、", "。", "，", ",", " ", "\t")
        parts = []
        buf = []
        in_quote = False
        for ch in text:
            if ch == '"':
                in_quote = not in_quote
                buf.append(ch)
            elif ch in seps and not in_quote:
                parts.append("".join(buf).strip())
                buf = []
            else:
                buf.append(ch)
        parts.append("".join(buf).strip())
        out = []
        for p in parts:
            p = p.strip()
            if p.startswith('"') and p.endswith('"') and len(p) >= 2:
                p = p[1:-1].strip()
            if p:
                out.append(p)
        return out

    def _update_map_file_label(self):
        if self._map_linked_file:
            self.map_file_label.setText("已关联: " + self._map_linked_file)
            self.map_link_btn.setToolTip(self._map_linked_file)
        else:
            self.map_file_label.setText("未关联映射表文件")
            self.map_link_btn.setToolTip("关联一个映射表文件路径")

    def _map_file_dir(self):
        cur = self._map_linked_file
        if cur and os.path.isdir(os.path.dirname(cur)):
            return os.path.dirname(cur)
        return ""

    def _on_link_mapping_file(self):
        """关联一个映射表文件（.xlsx/.csv），保存到 global settings 的 mapping_file。"""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择映射表文件",
            self._map_file_dir(),
            "映射表 (*.xlsx *.xls *.csv);;所有文件 (*.*)",
        )
        if not path:
            return
        self._map_linked_file = os.path.abspath(path)
        self._update_map_file_label()
        logger.info(f"已关联全量默认映射表文件: {self._map_linked_file}")
        self._on_debounced_change()  # 持久化 mapping_file

    def _on_open_mapping_file(self):
        """用系统默认编辑器（Excel/WPS）打开已关联的映射表文件。"""
        path = self._map_linked_file
        if not path or not os.path.isfile(path):
            start_dir = self._map_file_dir()
            picked, _ = QFileDialog.getOpenFileName(
                self,
                "选择要打开的映射表文件",
                start_dir,
                "映射表 (*.xlsx *.xls *.csv);;所有文件 (*.*)",
            )
            if not picked:
                return
            path = os.path.abspath(picked)
            self._map_linked_file = path
            self._update_map_file_label()
            self._on_debounced_change()
        try:
            if hasattr(os, "startfile"):
                os.startfile(path)
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(path))
            logger.info(f"已用系统默认程序打开全量映射表: {path}")
        except OSError as exc:
            QMessageBox.warning(
                self, "打开失败",
                f"无法打开映射表文件:\n{path}\n\n错误: {exc}",
            )

    def _on_import_mapping_file(self):
        """从 .xlsx/.csv 导入映射关系，合并进全量默认映射列表。

        两列结构（与「数据发送」映射表一致）：
          第一列 = 来源名（可 / ; 分隔多个，共享同一组接收人）
          第二列 = 接收人（可 / ; 分隔多人）
        合并规则：同一来源若已存在于文本里，用文件数据覆盖；文件里多行同源合并。
        """
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择映射表文件",
            self._map_file_dir(),
            "映射表 (*.xlsx *.xls *.csv);;所有文件 (*.*)",
        )
        if not path:
            return

        try:
            from modules.wps_cloud import read_one_sheet, read_sheet_names
        except ImportError as exc:
            QMessageBox.critical(self, "导入失败", f"缺少必要模块: {exc}")
            return

        try:
            names = read_sheet_names(path, log_fn=lambda m: logger.info(m))
            if not names:
                QMessageBox.warning(self, "导入失败", "无法读取工作表名称")
                return
            rows = read_one_sheet(path, names[0], log_fn=lambda m: logger.info(m))
        except Exception as exc:
            QMessageBox.warning(self, "导入失败", f"读取映射表失败:\n{path}\n\n错误: {exc}")
            return
        if not rows:
            QMessageBox.warning(self, "导入失败", "映射表为空或无有效数据")
            return

        # 先读回当前文本编辑框里的映射，合并后覆盖写回
        merged = {}
        for m in self._collect_mapping_rows_from_edit():
            merged[m["source_value"]] = m["recipients"]

        imported_count = 0
        for row in rows:
            if not row or len(row) < 2:
                continue
            srcs = self._plain_split_values(row[0])
            recips = self._plain_split_values(row[1])
            if not srcs or not recips:
                continue
            for s in srcs:
                merged[s] = list(recips)
                imported_count += 1

        lines = []
        for src, recips in merged.items():
            lines.append(
                f"{self._fmt_mapping_name(src)}="
                f"{';'.join(self._fmt_mapping_name(r) for r in recips)}"
            )
        self.map_edit.setPlainText("\n".join(lines))

        self._map_linked_file = os.path.abspath(path)
        self._update_map_file_label()
        self._on_debounced_change()
        logger.info(
            f"导入全量映射完成：合并后共 {len(merged)} 条映射"
            f"（本次新增/覆盖 {imported_count} 条），来源: {path}"
        )

    def _collect_mapping_rows_from_edit(self):
        """把全量映射文本解析为 [{source_value, recipients}]，key 去重合并。"""
        merged = {}
        for line in self.map_edit.toPlainText().splitlines():
            line = line.strip()
            if not line or "=" not in line:
                continue
            src_raw, recips_raw = line.split("=", 1)
            srcs = [s for s in self._split_mapping_names(src_raw) if s]
            recips = [r for r in self._split_mapping_names(recips_raw) if r]
            if not srcs or not recips:
                continue
            for src in srcs:
                merged.setdefault(src, []).extend(recips)
        return [
            {"source_value": src, "recipients": recips}
            for src, recips in merged.items()
        ]

    def _collect_settings(self):
        profile = str(self.rhythm_profile_combo.currentData() or "off")
        min_delay = round(float(self.rhythm_min_spin.value()), 1)
        max_delay = round(float(self.rhythm_max_spin.value()), 1)
        if min_delay > max_delay:
            min_delay, max_delay = max_delay, min_delay
            self.rhythm_min_spin.setValue(min_delay)
            self.rhythm_max_spin.setValue(max_delay)

        # 解析映射文本：来源1;来源2=接收人1;接收人2
        # - 等号左侧可含多个来源名（; 分隔），每个都映射到右侧整个接收人列表
        #   （"两个名字指向一个发送人" / 也可指向多个）
        # - 右侧接收人用 ; ； 分隔；"" 包裹的分隔符不生效
        # - 同一来源出现在多行时合并接收人，不做去重删除
        mappings = self._collect_mapping_rows_from_edit()
        return {
            "rhythm": {
                "enabled": self.rhythm_enable_check.isChecked(),
                "profile": profile,
                "min_delay": min_delay,
                "max_delay": max_delay,
            },
            "default_mapping": {
                "enabled": self.map_enable_check.isChecked(),
                "default_recipient": self.map_default_edit.text().strip(),
                "mapping_file": self._map_linked_file,
                "mappings": mappings,
            },
        }

    # ---------------- 落盘 + 应用 -----------------
    def _save(self):
        try:
            self.settings = self.config_manager.save_global_settings(self._collect_settings())
        except Exception as exc:
            logger.warning(f"保存全量设置失败: {exc}")

    def apply_rhythm(self):
        """将当前拟人节流设置应用到第三方驱动（startup 与变更时调用）。"""
        try:
            from modules.rhythm_settings import apply_rhythm
            apply_rhythm(self._collect_settings().get("rhythm") or {})
        except Exception as exc:
            logger.warning(f"应用拟人节流失败: {exc}")

    def get_default_mapping(self):
        """返回全量默认映射 (enabled, value_to_recipients, default_recipient)。"""
        enabled = self.map_enable_check.isChecked()
        if not enabled:
            return False, {}, ""
        value_to_recipients = {}
        for m in self._collect_settings().get("default_mapping", {}).get("mappings", []):
            value_to_recipients[m["source_value"]] = m["recipients"]
        default = self.map_default_edit.text().strip()
        return enabled, value_to_recipients, default

    def _on_immediate_change(self, *_):
        if self._loading:
            return
        self._save()
        self.apply_rhythm()
        self.settings_changed.emit()

    def _on_debounced_change(self, *_):
        if self._loading:
            return
        self._change_timer.start()

    def _flush_and_emit(self):
        if self._loading:
            return
        self._save()
        self.settings_changed.emit()


class MainWindow(QMainWindow):
    # 跨线程日志投递：子线程 emit -> 主线程 slot 写 UI
    _schedule_log_signal = pyqtSignal(str)
    # 跨线程触发投递：调度器线程 emit -> 主线程创建 QThread（确保 affinity 正确）
    _schedule_trigger_signal = pyqtSignal(object, str)
    # 发送进度开始：attach 可能发生在调度线程，用信号排队到主线程再碰 GUI/COM
    _progress_begin_signal = pyqtSignal()
    # 定时任务执行完毕（成功/超时/异常）后强制恢复标题/任务栏，
    # 不依赖 worker 的 finished_with_result 信号是否正确到达，
    # 避免无人值守时窗口标题永久停在“发送中…”
    # 参数：失败数量, 成功数量, 失败人名列表（供兜底补发完成 Toast）
    _schedule_done_signal = pyqtSignal(int, int, list)
    # 定时任务异常（超时强制停止等）需要弹常驻通知时，从调度线程
    # 经此信号投递到主线程弹 Toast（msg, success）
    _schedule_toast_signal = pyqtSignal(str, bool)

    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setGeometry(50, 50, 850, 580)
        # 弹性窗口最小地板：折叠后可缩到这个尺寸，不至于过小
        self.setMinimumSize(520, 300)
        self._closing_requested = False
        self._close_ready = False
        self._shutdown_poll_count = 0
        self._tray_hint_shown = False
        # 窗口弹性自适应状态：首次显示后按内容尺寸适配一次；几何动画句柄
        self._startup_fit_done = False
        self._geo_anim: Optional[QPropertyAnimation] = None

        self.schedule_store = ScheduleStore()
        self.schedule_dispatcher = ScheduleDispatcher(self.schedule_store)
        self.schedule_dispatcher.add_log_handler(self._schedule_dispatch_log)
        self.schedule_dispatcher.add_handler(self._on_schedule_triggered)
        # 定时触发时并发串行化（避免同时多个任务抢微信窗口）
        self._schedule_running_lock = threading.Lock()
        self._schedule_worker = None
        # 每次定时任务执行周期内"完成 Toast"只允许弹一次：
        # 正常路径（_on_schedule_result）与兜底路径（_schedule_done_signal）
        # 都会尝试弹，由该标志去重；新任务在主线程触发时复位
        self._schedule_result_toast_shown = False

        # 关键：跨线程日志用 pyqtSignal 投递，不能用 QTimer.singleShot(0,...)
        # 因为子线程没有 Qt event loop，singleShot 永远不会触发
        self._schedule_log_signal.connect(self._on_schedule_log_signal)
        self._schedule_trigger_signal.connect(self._on_schedule_trigger_in_main)
        # 进度开始信号始终在主线程执行（receiver=MainWindow），保证 setWindowTitle/
        # winId()/COM/托盘 全部在主线程，避免调度线程跨线程操作 GUI
        self._progress_begin_signal.connect(self._begin_send_progress)
        # 定时任务结束后强制恢复进度（兜底，不依赖 worker 结束信号），
        # 并在正常 Toast 信号意外丢失时补发完成通知
        self._schedule_done_signal.connect(self._on_schedule_done_fallback)
        # 定时任务异常通知（超时强制停止等）在主线程弹 Toast
        self._schedule_toast_signal.connect(self._show_schedule_toast)

        # 电脑锁定守护：阻止定时发送期间电脑自动锁定，任务完成后自动锁回
        self.workstation_guard = WorkstationGuard(
            log_fn=self._post_schedule_log_from_worker
        )

        self.init_ui()
        self.init_tray()

        # ---- 在线更新状态 ----
        # 首次运行检查一次 + 长运行每满 24 小时检查一次；
        # 同一新版本自动提醒只弹一次，用户可跳过；手动检查不受限制
        self._update_state = UpdateState.load()
        self._update_check_worker: Optional[UpdateCheckWorker] = None
        self._update_dialog: Optional[UpdateDialog] = None
        self._update_auto_timer = QTimer(self)
        self._update_auto_timer.setInterval(updater.AUTO_TICK_MS)
        self._update_auto_timer.timeout.connect(self._on_update_timer_tick)
        # 安全护栏：offscreen（自动化测试/无头环境）绝不启动真实调度器，
        # 否则测试进程会加载并 mark_fired 用户真实的定时任务，导致当天任务被抢占
        if os.environ.get("QT_QPA_PLATFORM", "").lower() == "offscreen":
            self.schedule_tab.dispatcher_status_label.setText(
                "调度器状态: 测试模式未启动"
            )
        else:
            self.schedule_dispatcher.start()
            self.schedule_tab.dispatcher_status_label.setText("调度器状态: 运行中")

    def init_ui(self):
        self.table_filter_tab = TableFilterTab()
        self.schedule_tab = ScheduleTab(self.schedule_store, self.schedule_dispatcher)
        # 让 ScheduleTab 日志同时写入原表格发送主界面的日志面板，保持统一查看
        self.schedule_tab.log_callback = self._schedule_log_to_main
        self.monitor_store = MonitorManager()
        self.monitor_tab = MonitorTab(self.monitor_store)
        # 让监控日志同时写入原表格发送主界面的日志面板，保持统一查看
        self.monitor_tab.log_callback = self._schedule_log_to_main

        self.tab_widget = _ElasticTabWidget()
        # 用 ElasticPage 包装：非当前页不报告尺寸提示，窗口才能弹性收缩
        page1 = ElasticPage()
        pl1 = QVBoxLayout(page1)
        pl1.setContentsMargins(0, 0, 0, 0)
        pl1.addWidget(self.table_filter_tab)
        page2 = ElasticPage()
        pl2 = QVBoxLayout(page2)
        pl2.setContentsMargins(0, 0, 0, 0)
        pl2.addWidget(self.schedule_tab)
        page3 = ElasticPage()
        pl3 = QVBoxLayout(page3)
        pl3.setContentsMargins(0, 0, 0, 0)
        pl3.addWidget(self.monitor_tab)
        # 全量设置：应用级全局设置（拟人节流 + 全量默认映射）
        self.config_manager = ConfigManager()
        self.global_settings_tab = GlobalSettingsTab(self.config_manager)
        page4 = ElasticPage()
        pl4 = QVBoxLayout(page4)
        pl4.setContentsMargins(0, 0, 0, 0)
        pl4.addWidget(self.global_settings_tab)
        self.tab_widget.addTab(page1, "📊 数据发送")
        self.tab_widget.addTab(page2, "⏰ 定时发送")
        self.tab_widget.addTab(page3, "📡 监控")
        self.tab_widget.addTab(page4, "⚙️ 全量设置")

        # 全局精简/详细开关：放 QTabWidget 右上角 corner
        self._compact_mode = False
        self.compact_btn = QPushButton("📑 精简模式")
        self.compact_btn.setStyleSheet(_compact_btn_css())
        self.compact_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.compact_btn.clicked.connect(self._on_compact_toggle)
        self.tab_widget.setCornerWidget(self.compact_btn)
        self.tab_widget.currentChanged.connect(self._sync_compact_btn)

        self.setCentralWidget(self.tab_widget)

        # 注入 worker 创建回调：任何发送 worker 启动后，在任务栏/窗口标题显示进度，
        # 主窗口被微信窗口挡住时也能在任务栏看到发送进度
        self.table_filter_tab.worker_created_cb = self._attach_global_progress_table
        self.schedule_tab.worker_created_cb = self._attach_global_progress_schedule
        self.table_filter_tab.progress_dismissed_cb = self._dismiss_send_progress
        self._taskbar = None
        self._send_progress_active = False
        # 门闩：数据发送完成弹窗等待用户点击期间，禁止 worker.finished 提前清除任务栏
        self._progress_waiting_dismiss = False

    def _current_tab_content(self) -> QWidget:
        """取当前 tab 内的真实业务页面（穿透 ElasticPage 包装层）。"""
        page = self.tab_widget.currentWidget()
        if isinstance(page, ElasticPage):
            return page.content_widget()
        return page

    def _on_compact_toggle(self):
        """全局精简/详细模式：折叠/展开当前 tab 的所有可折叠分组。"""
        self._compact_mode = not self._compact_mode
        widget = self._current_tab_content()
        groups = getattr(widget, "_collapsible_groups", [])
        for g in groups:
            g.set_collapsed(self._compact_mode)
        self.compact_btn.setText("📄 详细模式" if self._compact_mode else "📑 精简模式")

    def _sync_compact_btn(self, _index):
        """切换 tab 时根据当前 tab 的折叠状态更新按钮文字。"""
        widget = self._current_tab_content()
        groups = getattr(widget, "_collapsible_groups", [])
        if not groups:
            return
        all_collapsed = all(g.is_collapsed() for g in groups)
        self._compact_mode = all_collapsed
        self.compact_btn.setText("📄 详细模式" if all_collapsed else "📑 精简模式")
        # 切换 tab 后按新页理想尺寸适应（折叠的分组不计入 sizeHint，
        # 已展开内容完整显示，不再出现定时任务页文字被压扁）
        QTimer.singleShot(0, lambda: self.fit_to_content(True))

    def changeEvent(self, event):
        """系统明暗（夜间/日间）主题切换时，重刷精简按钮配色。"""
        super().changeEvent(event)
        if event.type() == QEvent.Type.PaletteChange:
            try:
                self.compact_btn.setStyleSheet(_compact_btn_css())
            except Exception:
                pass
        elif event.type() == QEvent.Type.WindowStateChange:
            # 从最小化恢复：最小化期间折叠/展开被跳过，恢复后补一次自适应
            if not self.isMinimized() and not self.isMaximized() \
                    and not self.isFullScreen() and self._startup_fit_done:
                QTimer.singleShot(0, lambda: self.fit_to_content(True))

    def showEvent(self, event):
        """首次显示后按内容理想尺寸自适应一次，修复启动时部分文字被压缩。"""
        super().showEvent(event)
        if not self._startup_fit_done:
            self._startup_fit_done = True
            QTimer.singleShot(0, lambda: self.fit_to_content(True))
            # 首次运行：延迟后自动检查一次更新（避开启动高峰）；
            # offscreen 测试环境不发起真实网络请求
            if os.environ.get("QT_QPA_PLATFORM", "").lower() != "offscreen":
                QTimer.singleShot(
                    updater.STARTUP_CHECK_DELAY_MS,
                    lambda: self._start_update_check(manual=False))
                self._update_auto_timer.start()

    def _animate_window_geometry(self, target: QRect) -> None:
        """平滑动画过渡到目标窗口几何（折叠/展开时的果冻弹性跟随）。"""
        # 目标几何与当前一致：停掉可能在跑的旧动画，无需新建动画
        if self.geometry() == target:
            if (
                self._geo_anim is not None
                and self._geo_anim.state() == QAbstractAnimation.State.Running
            ):
                try:
                    self._geo_anim.stop()
                except Exception:
                    pass
            return
        if (
            self._geo_anim is not None
            and self._geo_anim.state() == QAbstractAnimation.State.Running
            and self._geo_anim.endValue() == target
        ):
            return  # 目标相同且动画进行中，无需重启
        if self._geo_anim is not None:
            try:
                self._geo_anim.stop()
            except Exception:
                pass
            self._geo_anim.deleteLater()
        anim = QPropertyAnimation(self, b"geometry", self)
        anim.setDuration(300)
        anim.setEasingCurve(QEasingCurve.Type.OutQuart)
        anim.setEndValue(target)
        self._geo_anim = anim
        anim.start()

    def fit_to_content(self, expanding=None, animate_window=False):
        """弹性窗口：折叠/展开分组后，窗口尺寸跟随内容自适应收缩或放大。

        - expanding=False（折叠触发）：目标取当前页 minimumSizeHint，窗口收缩到
          刚好不裁切内容的最小尺寸；
        - expanding=True（展开触发）：目标取当前页 sizeHint，窗口放大到理想尺寸；
        - expanding=None（如切换 tab）：仅保证不小于 minimumSizeHint，不强行放大。
        - 注意不能用 QTabWidget.sizeHint()：QStackedWidget 会取所有页面的最大值，
          当前页折叠后它仍报其它页的大尺寸，无法用于弹性收缩。
        - 限制在屏幕可用区域内，保留最小尺寸地板；最大化/全屏/未显示时不干预。
        """
        try:
            if (
                not self.isVisible()
                or self.isMinimized()
                or self.isMaximized()
                or self.isFullScreen()
            ):
                return
            # 关键：先同步分发挂起的布局事件，刷新各 widget 的 sizeHint/minimumSizeHint
            # 缓存（QTimer(0) 回调可能先于 LayoutRequest 执行，否则会读到折叠前的旧尺寸）
            app = QApplication.instance()
            if app is not None:
                app.sendPostedEvents(None, QEvent.Type.LayoutRequest)
            page = self.tab_widget.currentWidget()
            if page is None:
                return
            min_hint = page.minimumSizeHint()
            if not min_hint.isValid() or min_hint.width() <= 0:
                min_hint = page.sizeHint()
            if expanding:
                hint = page.sizeHint()
                if not hint.isValid() or hint.width() <= 0:
                    hint = min_hint
                # 展开时理想尺寸与最小尺寸取较大者
                target_w = max(hint.width(), min_hint.width())
                target_h = max(hint.height(), min_hint.height())
            else:
                # expanding=False（折叠）或 None（切换 tab）：
                # 窗口收缩/适应到内容最小尺寸，不浪费空间
                target_w, target_h = min_hint.width(), min_hint.height()
            # 加上非页面部分增量（tab 栏高度、边框等）：当前窗口尺寸减去页面实际尺寸
            delta_w = self.width() - page.width()
            delta_h = self.height() - page.height()
            target_w += max(0, delta_w)
            target_h += max(0, delta_h)

            screen = (
                QGuiApplication.screenAt(self.geometry().center())
                or QGuiApplication.primaryScreen()
            )
            if screen is None:
                return
            avail = screen.availableGeometry()
            min_w, min_h = 520, 300
            w = max(min_w, min(target_w, avail.width()))
            h = max(min_h, min(target_h, avail.height()))
            # 关键：setGeometry 设的是客户区几何，屏幕可用区约束的是整个窗口
            # 框架（含标题栏）。必须先把框架四边的开销（标题栏高度、边框宽度）
            # 扣除，否则内容接近满屏时窗口框架顶端会被顶出屏幕外，
            # 表现为"标题栏和最小化/最大化/关闭按钮不见了"。
            frame = self.frameGeometry()
            geo = self.geometry()
            # Qt 框架几何包围客户区：标题栏在客户区上方（frame.top < geo.top），
            # 故各方向开销为：上 = geo.top-frame.top，下/右 = frame-geo
            ft = max(0, geo.top() - frame.top())         # 标题栏高度
            fl = max(0, geo.left() - frame.left())       # 左边框宽度
            fr = max(0, frame.right() - geo.right())     # 右边框宽度
            fb = max(0, frame.bottom() - geo.bottom())   # 底边框宽度
            w = max(min_w, min(target_w, avail.width() - fl - fr))
            h = max(min_h, min(target_h, avail.height() - ft - fb))
            # 保持整个窗口框架（含标题栏）在屏幕可用区内
            x = max(avail.left() + fl, min(self.x(), avail.left() + avail.width() - fr - w))
            y = max(avail.top() + ft, min(self.y(), avail.top() + avail.height() - fb - h))
            if animate_window:
                # 果冻弹性：窗口平滑动画过渡到目标几何
                self._animate_window_geometry(QRect(x, y, w, h))
            else:
                # 立即适配：停掉进行中的几何动画，避免旧动画回写几何值
                if self._geo_anim is not None:
                    try:
                        self._geo_anim.stop()
                    except Exception:
                        pass
                    self._geo_anim.deleteLater()
                    self._geo_anim = None
                self.setGeometry(x, y, w, h)
        except Exception:
            pass

    # ------------------------- 全局发送进度（任务栏/标题/托盘） -------------------------
    def _ensure_taskbar(self):
        """懒加载任务栏进度条（需要 native 窗口 HWND，show 之后才有效）。"""
        if self._taskbar is not None:
            return self._taskbar
        hwnd = int(self.winId()) if self.winId() else 0
        try:
            from modules.taskbar_progress import TaskbarProgress
            tb = TaskbarProgress(hwnd)
            if tb.available:
                self._taskbar = tb
                self.log(f"[系统] 任务栏进度条已就绪(hwnd=0x{hwnd:X})")
            else:
                self._taskbar = False
                self.log("[系统] 任务栏进度条不可用(ITaskbarList3 初始化失败，已静默降级)")
        except Exception as exc:
            self._taskbar = False
            self.log(f"[系统] 任务栏进度条初始化异常: {exc}")
        return self._taskbar if self._taskbar is not False else None

    def _begin_send_progress(self):
        # 终态门禁：完成/出错弹窗正等待用户点击时，忽略新的“开始”（如定时任务
        # 恰好并发触发），不得把定格标题/任务栏冲回“发送中…”。新一轮正常发送
        # 只会在弹窗关闭（门闩复位）后由用户手动启动。
        if self._progress_waiting_dismiss:
            return
        self._send_progress_active = True
        tb = self._ensure_taskbar()
        if tb:
            tb.set_indeterminate()
            self.log("[系统] 任务栏进度条：开始显示发送进度")
        else:
            self.log("[系统] 任务栏进度条未就绪，仅在窗口标题和托盘显示进度")
        self.setWindowTitle(f"📤 发送中… - {APP_TITLE}")
        try:
            self.tray_icon.setToolTip(f"{APP_TITLE}\n📤 发送中…")
        except Exception:
            pass

    def _update_send_progress(self, current, total):
        # 终态门禁：定格期间迟到的 progress（队列残留/并发 worker）一律丢弃
        if self._progress_waiting_dismiss:
            return
        if not self._send_progress_active:
            self._begin_send_progress()
        if self._progress_waiting_dismiss:
            return
        try:
            current = max(0, int(current))
            total = max(1, int(total))
            pct = int(current * 100 / total)
        except Exception:
            return
        tb = self._ensure_taskbar()
        if tb:
            tb.set_value(current, total)
        title = f"📤 发送中 {current}/{total} ({pct}%) - {APP_TITLE}"
        self.setWindowTitle(title)
        try:
            self.tray_icon.setToolTip(f"{APP_TITLE}\n{title}")
        except Exception:
            pass

    def _update_schedule_send_progress(self, p_cur, p_total, s_cur, s_total):
        """定时配置链发送进度：标题同时显示发送进度和链路进度。

        格式：📤 发送中 5/10 (50%) 链路1/3 - {APP_TITLE}
        """
        # 终态门禁：定格期间迟到的 progress 一律丢弃
        if self._progress_waiting_dismiss:
            return
        if not self._send_progress_active:
            self._begin_send_progress()
        if self._progress_waiting_dismiss:
            return
        try:
            s_cur = max(0, int(s_cur))
            s_total = max(1, int(s_total))
            p_cur = max(0, int(p_cur))
            p_total = max(1, int(p_total))
            pct = int(s_cur * 100 / s_total)
        except Exception:
            return
        tb = self._ensure_taskbar()
        if tb:
            tb.set_value(s_cur, s_total)
        title = (f"📤 发送中 {s_cur}/{s_total} ({pct}%) "
                 f"链路{p_cur}/{p_total} - {APP_TITLE}")
        self.setWindowTitle(title)
        try:
            self.tray_icon.setToolTip(f"{APP_TITLE}\n{title}")
        except Exception:
            pass

    def _show_send_done(self, success, failed, total):
        """发送结束的“定格”展示：任务栏满进度（有失败变红），标题显示已发送。

        该状态一直保持到用户点掉完成弹窗（_dismiss_send_progress），
        不再像之前那样 result 一到就清除，导致弹窗还开着任务栏已无数。
        """
        if not self._send_progress_active:
            self._begin_send_progress()
        self._progress_waiting_dismiss = True
        # 看门狗：完成弹窗若因窗口在托盘/系统异常/无人值守长时间未关闭，
        # 60 秒后自动收尾，杜绝标题永久停在发送态
        QTimer.singleShot(60000, self._dismiss_progress_if_waiting)
        try:
            success = max(0, int(success))
            failed = max(0, int(failed))
            total = max(1, int(total))
        except Exception:
            success, failed, total = 0, 0, 1
        done = min(success + failed, total)
        tb = self._ensure_taskbar()
        if tb:
            tb.set_value(done, total)
            if failed:
                tb.set_error()
        if failed:
            title = f"⚠ 发送完成 成功{success} 失败{failed}（共{total}）- {APP_TITLE}"
        else:
            title = f"✅ 已发送完成 共{total}人 - {APP_TITLE}"
        self.setWindowTitle(title)
        try:
            self.tray_icon.setToolTip(f"{APP_TITLE}\n{title}")
        except Exception:
            pass

    def _show_send_error(self):
        """发送异常的“定格”展示：红色任务栏 + 标题提示，等用户点掉错误弹窗。"""
        if not self._send_progress_active:
            self._begin_send_progress()
        self._progress_waiting_dismiss = True
        # 看门狗：错误弹窗同样可能无人关闭，60 秒后自动收尾
        QTimer.singleShot(60000, self._dismiss_progress_if_waiting)
        tb = self._ensure_taskbar()
        if tb:
            tb.set_error()
        title = f"❌ 发送出错 - {APP_TITLE}"
        self.setWindowTitle(title)
        try:
            self.tray_icon.setToolTip(f"{APP_TITLE}\n{title}")
        except Exception:
            pass

    def _dismiss_send_progress(self, failed=0):
        """用户在完成/错误弹窗点击确定后调用：恢复标题、清除任务栏进度。"""
        self._progress_waiting_dismiss = False
        self._finish_send_progress(failed)

    def _dismiss_progress_if_waiting(self):
        """看门狗兜底：弹窗超时未关闭时自动恢复标题/任务栏。"""
        if self._progress_waiting_dismiss:
            self._progress_waiting_dismiss = False
            self._finish_send_progress(0)

    def _finish_send_progress(self, failed=0):
        # 门闩保护：完成弹窗仍在等待点击时（可能是并发的定时任务结束来清理），
        # 不得提前清除定格状态；弹窗关闭时 _dismiss_send_progress 会统一收尾
        if self._progress_waiting_dismiss:
            return
        tb = self._ensure_taskbar()
        if tb and failed:
            tb.set_error()
        elif tb:
            tb.clear()
        self.setWindowTitle(APP_TITLE)
        try:
            self.tray_icon.setToolTip(APP_TITLE)
        except Exception:
            pass
        self._send_progress_active = False
        # 有失败时红色状态停留 4 秒再清，提示用户注意
        if tb and failed:
            QTimer.singleShot(4000, lambda: tb.clear() if tb else None)

    def _force_finish_schedule_progress(self, failed=0):
        """定时任务结束后强制恢复标题/任务栏（兜底机制）。

        不看门闩、不看 _send_progress_active，无条件清除，确保无人值守时
        窗口标题和任务栏进度条不会永久停在“发送中…”。
        _on_schedule_result 已做过正常收尾时这里是空操作，无副作用。
        """
        self._progress_waiting_dismiss = False
        tb = self._ensure_taskbar()
        if tb:
            try:
                tb.clear()
            except Exception:
                pass
        try:
            self.setWindowTitle(APP_TITLE)
        except Exception:
            pass
        try:
            self.tray_icon.setToolTip(APP_TITLE)
        except Exception:
            pass
        self._send_progress_active = False
        # 有失败时自动展开日志面板，确保用户能看到失败详情
        if failed:
            self._expand_log_on_failure()

    def _show_schedule_toast(self, msg, success):
        """定时任务异常通知（超时强制停止等），主线程弹 Toast。
        success=False 时常驻等待手动关闭（ToastNotification 内强制）。"""
        try:
            ToastNotification.show_toast(
                "定时任务", str(msg), success=bool(success),
                duration=8000, parent=self)
        except Exception:
            pass

    def _expand_log_on_failure(self):
        """发送有失败时自动展开对应日志面板，确保用户能看到结果。

        手动发送/auto_send → 展开数据发送页日志面板；
        定时任务/立即执行/配置链 → 展开定时页日志面板；
        两个都展开以防当前不在对应 tab。
        所有操作都在主线程（由信号排队触发），线程安全。
        """
        try:
            tft = self.table_filter_tab
            if tft is not None and hasattr(tft, "log_group"):
                lg = tft.log_group
                if lg is not None and lg.is_collapsed():
                    lg.set_collapsed(False, animate=True)
        except Exception:
            pass
        try:
            st = self.schedule_tab
            if st is not None and hasattr(st, "schedule_log_group"):
                lg = st.schedule_log_group
                if lg is not None and lg.is_collapsed():
                    lg.set_collapsed(False, animate=True)
        except Exception:
            pass

    def _attach_global_progress_table(self, worker):
        """数据发送 worker（进度信号为 0-100 百分比）。

        注意：这些信号由 worker 线程发出，槽必须是 MainWindow 的绑定方法
        （receiver 为主线程 QObject，AutoConnection 自动变 QueuedConnection，
        在主线程执行）。不能用 lambda——lambda 无 receiver，会在 worker 线程
        直接执行，跨线程操作 GUI/COM 会导致界面卡死。
        """
        try:
            self._progress_begin_signal.emit()
            worker.signals.progress.connect(self._on_table_progress_pct)
            # result = (success, failed, total, failed_tasks)
            worker.signals.result.connect(self._on_table_result)
            worker.signals.error.connect(self._on_table_error)
            worker.finished.connect(self._clear_progress_if_idle)
        except Exception:
            pass

    def _on_table_progress_pct(self, current, total):
        self._update_send_progress(current, total)

    def _on_table_result(self, r):
        # 不立即清除：定格在“已发送完成”状态，保持到用户点掉完成弹窗
        if isinstance(r, (list, tuple)) and len(r) >= 3:
            success, failed, total = r[0], r[1], r[2]
        else:
            success, failed, total = 0, 0, 1
        self._show_send_done(success, failed, total)

    def _on_table_error(self, _e):
        # 不立即清除：定格在出错红色状态，保持到用户点掉错误弹窗
        self._show_send_error()

    def _attach_global_progress_schedule(self, worker):
        """定时发送 worker（进度信号为 current/total）。

        同理，finished_with_result 虽在另一处用 DirectConnection 连接，但本
        连接是独立的；绑定方法 receiver 在主线程，AutoConnection → 排队到主线程。
        """
        try:
            # emit 线程安全；_begin_send_progress 经排队在主线程执行，
            # 保证本方法即便被调度线程调用也不会跨线程碰 GUI/COM
            self._progress_begin_signal.emit()
            worker.progress.connect(self._update_send_progress)
            # 配置链特有的发送进度信号（链路当前, 链路总数, 发送当前, 发送总数）
            if hasattr(worker, 'send_progress'):
                worker.send_progress.connect(self._update_schedule_send_progress)
            worker.finished_with_result.connect(self._on_schedule_result)
            # 兜底：若结束信号因任何原因未到达，QThread 结束时也恢复标题/清任务栏，
            # 避免窗口标题永久停留在“发送中…”
            worker.finished.connect(self._clear_progress_if_idle)
        except Exception:
            pass

    def _show_schedule_completion_toast(self, ok_n, fail, recipients):
        """构造并弹出"定时任务完成"Toast；每个执行周期恰好一次（守卫去重）。

        正常路径与 _schedule_done_signal 兜底路径都会调用本方法。
        """
        try:
            if self._schedule_result_toast_shown:
                return
            self._schedule_result_toast_shown = True
            ok_n = int(ok_n) if isinstance(ok_n, (int, float)) else 0
            fail = int(fail) if isinstance(fail, (int, float)) else 0
            if fail == 0:
                toast_msg = f"定时任务发送成功，共 {ok_n} 条"
            else:
                names = [str(n) for n in (recipients or []) if str(n).strip()]
                toast_msg = f"定时任务完成：成功 {ok_n}，失败 {fail}"
                if names:
                    shown = names[:5]
                    suffix = f" 等{len(names)}人" if len(names) > 5 else ""
                    toast_msg += f"\n失败：{'、'.join(shown)}{suffix}"
                toast_msg += "（点击查看）"
            logger.info(
                "[定时] 弹出完成通知：成功=%s 失败=%s 失败人数=%s",
                ok_n, fail, len(recipients or []))
            ToastNotification.show_toast(
                "定时任务完成", toast_msg,
                success=(fail == 0),
                duration=8000, parent=self)
        except Exception:
            logger.exception("[定时] 完成通知弹出失败")

    def _on_schedule_result(self, _ok, fail, _recipients):
        # 定时任务（纯文字 / 配置链）没有完成确认弹窗，标题和任务栏必须在这里
        # 无条件恢复：即便残留了手动发送的"等待关闭弹窗"门闩，也不能挡住定时收尾，
        # 否则无人值守时窗口标题会永久停在"发送中…"。
        self._progress_waiting_dismiss = False
        self._finish_send_progress(fail)
        logger.info(
            "[定时] worker 完成信号到达：成功=%s 失败=%s", _ok, fail)
        # Toast 通知（定时任务完成；有失败时常驻等待手动关闭，并列出失败人名）
        self._show_schedule_completion_toast(_ok, fail, _recipients)
        # 有失败时自动展开日志面板
        if fail:
            self._expand_log_on_failure()

    def _on_schedule_done_fallback(self, failed, success, failed_names):
        """_schedule_done_signal 的主线程槽：无条件收尾 + Toast 兜底。

        正常情况下 _on_schedule_result 已先到（同为排队事件，emit 更早），
        Toast 守卫会跳过这里的重复弹窗；若 worker 的 finished_with_result
        排队事件意外丢失，则由这里补发，保证无人值守时通知不缺席。
        """
        try:
            self._force_finish_schedule_progress(failed)
        except Exception:
            logger.exception("[定时] 兜底收尾异常")
        if not self._schedule_result_toast_shown:
            logger.warning(
                "[定时] worker 完成信号未到达，由 _schedule_done_signal 兜底弹通知")
            self._show_schedule_completion_toast(success, failed, failed_names)

    def _clear_progress_if_idle(self):
        # finished 在 result/error 之后触发，会借完成弹窗的局部事件循环投递到主线程。
        # 若完成弹窗仍在等待用户点击，必须保留“已发送完成”定格状态，禁止提前清除；
        # 仅在 result/error 都没到达的异常退出场景下做兜底清除。
        if self._progress_waiting_dismiss:
            return
        if self._send_progress_active:
            self._finish_send_progress(0)

    def _schedule_log_to_main(self, message):
        try:
            self.table_filter_tab.log(message)
        except Exception:
            pass

    def _is_other_send_running(self, exclude=None):
        """是否有其他微信发送 worker 正在运行（手动表格发送/立即执行/定时触发）。

        所有发送共用 WeChatSender 单例 + wxauto GUI 自动化，两个 worker
        并发会争抢微信窗口导致发错对象。QThread.isRunning() 可跨线程安全读取。
        只做线程安全的状态读取，不碰任何 QWidget。
        """
        candidates = []
        try:
            candidates.append(self.table_filter_tab.worker)
        except Exception:
            pass
        try:
            candidates.append(self.schedule_tab.send_worker)
        except Exception:
            pass
        candidates.append(getattr(self, "_schedule_worker", None))
        for worker in candidates:
            if worker is None or worker is exclude:
                continue
            try:
                if worker.isRunning():
                    return True
            except Exception:
                continue
        return False

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
        """主线程 slot：创建 worker（affinity = 主线程），再交给 ScheduleRun 线程执行。

        不再做"与其他发送任务冲突则跳过"的预检：全局发送批次锁会让本任务
        排队等待其他发送完成后再继续（排队无时间上限，用户确认）。
        """
        # 主线程内容校验（调度器已校验一次，这里双保险，防止数据在窗口内被改坏）
        problem = ScheduleDispatcher._validate_task_content(task)
        if problem:
            self._post_schedule_log_from_worker(
                f"[定时] ⚠ 任务「{task.name}」时间点 {slot} 跳过执行：{problem}"
            )
            return
        if getattr(task, "kind", "message") == "profiles":
            # 过滤掉已被移动/删除的配置：存在的照常链式执行，缺失的逐个告警
            valid_paths, missing_paths = [], []
            for p in (task.profile_paths or []):
                (valid_paths if os.path.isfile(str(p)) else missing_paths).append(str(p))
            for p in missing_paths:
                self._post_schedule_log_from_worker(
                    f"[定时] ⚠ 配置文件不存在，已跳过：{p}"
                )
            worker = ProfileChainWorker(
                profile_paths=valid_paths,
                log_callback=self._post_schedule_log_from_worker,
                minimize_after=getattr(task, "minimize_after", True),
            )
        else:
            worker = ScheduleSendWorker(
                recipients=list(task.recipients),
                message=task.message,
                chat_delay=task.chat_delay,
                send_interval=task.send_interval,
                log_callback=self._post_schedule_log_from_worker,
                default_city=task.default_city,
                minimize_after=getattr(task, "minimize_after", True),
                attachment=getattr(task, "attachment", "") or "",
                send_order=list(getattr(task, "send_order", None) or ["message"]),
            )
        self._schedule_worker = worker
        # 新执行周期：复位完成 Toast 去重守卫
        self._schedule_result_toast_shown = False
        logger.info(
            "[定时] 主线程创建 worker：任务=%s kind=%s 时间点=%s",
            getattr(task, "name", ""), getattr(task, "kind", "message"), slot)
        # 后台触发的任务（尤其配置链可能跑很久）：启用停止按钮，
        # worker 结束后（finished 在主线程排队执行）恢复按钮状态
        try:
            self.schedule_tab.stop_send_btn.setEnabled(True)
            worker.finished.connect(self._on_main_schedule_worker_finished_ui)
        except Exception:
            pass
        threading.Thread(
            target=self._run_schedule_task_serialized,
            args=(task, slot, worker),
            daemon=True,
            name=f"ScheduleRun-{task.id[:8]}",
        ).start()

    def _on_main_schedule_worker_finished_ui(self):
        """后台定时 worker 结束后恢复定时页按钮（主线程槽）。"""
        try:
            run_now_worker = self.schedule_tab.send_worker
            still_busy = bool(
                run_now_worker is not None and run_now_worker.isRunning()
            )
            self.schedule_tab.stop_send_btn.setEnabled(still_busy)
            self.schedule_tab.run_now_btn.setEnabled(True)
        except Exception:
            pass

    def _run_schedule_task_serialized(self, task: ScheduleTask, slot: str, worker: "ScheduleSendWorker"):
        with self._schedule_running_lock:
            # 与其他定时任务用本锁互斥；跨通道（手动/监控/立即执行）不再做
            # "冲突则跳过"，统一交给全局发送批次锁排队：本任务会在其他发送
            # 完成后自动继续（排队无时间上限）。worker 排队期间被手动停止
            # 时 acquire_batch 会返回 False，worker 自行结束并发出结束信号。
            # 按任务配置启动防锁定守护
            if getattr(task, "keep_unlocked", False):
                if not self.workstation_guard.is_running():
                    self.workstation_guard.start()
                    self._post_schedule_log_from_worker(
                        f"[定时] 任务「{task.name}」已开启防锁定守护"
                    )
            is_profile_chain = getattr(task, "kind", "message") == "profiles"
            if is_profile_chain:
                unit_count = len(task.profile_paths or [])
                self._post_schedule_log_from_worker(
                    f"[定时] 开始执行配置链任务「{task.name}」 时间点 {slot} "
                    f"配置数 {unit_count}"
                )
            else:
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
            # 任务栏/标题进度：默认连接（AutoConnection），槽在主线程执行，可安全更新 GUI
            try:
                self._attach_global_progress_schedule(worker)
            except Exception:
                pass
            worker.start()

            if is_profile_chain:
                # 配置链：每个配置可能含多人，按 30 分钟/个配置估超时
                total_timeout = max(300.0, 1800.0 * max(1, unit_count))
            else:
                total_timeout = max(60.0, 60.0 * (len(task.recipients) or 1) * 10.0)
            total_timeout = min(total_timeout, 12 * 3600.0)

            # 等待策略：先等 worker 拿到全局发送批次锁（排队阶段无超时上限，
            # 用户确认"排队无时间上限"）；拿到锁后才开始计算任务总超时。
            # 排队期间被手动停止时，worker 未拿锁即结束（finished 置位）。
            run_started_ts = None  # worker 拿到批次锁、真正开始发送的时间
            timed_out = False
            while not finished.is_set():
                if run_started_ts is None and getattr(worker, "_batch_acquired", False):
                    run_started_ts = time.time()
                if run_started_ts is not None:
                    if time.time() - run_started_ts >= total_timeout:
                        timed_out = True
                        break
                # 兜底：worker 线程异常退出且从未拿到锁，避免无限轮询
                try:
                    if not worker.isRunning() and run_started_ts is None:
                        break
                except Exception:
                    pass
                if finished.wait(1.0):
                    break
            if timed_out:
                self._post_schedule_log_from_worker(
                    f"[定时] ⚠ 任务「{task.name}」执行超时（{int(total_timeout)}秒），强制停止"
                )
                try:
                    worker.stop()
                except Exception:
                    pass
                # stop() 只置标志位，若 worker 正阻塞在 wxauto 调用里需要等其退出，
                # 否则本锁释放后排队的下一个定时任务会与 zombie worker 并发抢微信。
                # QThread.wait() 跨线程安全；最多再等 30 秒，到期也放锁避免死等。
                try:
                    worker.wait(30000)
                except Exception:
                    pass
                # 超时（卡死）时 worker 的 finished_with_result 不会到达，
                # _on_schedule_result 的完成通知也不会弹；这里补发常驻失败通知，
                # 避免无人值守时任务挂了用户毫无感知（2026-09-28 弹窗卡死事件）
                try:
                    still_alive = worker.isRunning()
                except Exception:
                    still_alive = True
                timeout_msg = (
                    f"任务「{task.name}」执行超时已强制停止"
                    f"（成功 {result['success']}，失败 {result['failed']}，"
                    f"其余未发送），请查看日志"
                )
                if still_alive:
                    timeout_msg += "；发送线程仍阻塞，建议重启程序"
                try:
                    self._schedule_toast_signal.emit(timeout_msg, False)
                    # 超时已弹专属常驻失败通知：阻止兜底路径再弹一个
                    # "定时任务完成"通知，避免同一任务双弹窗。
                    # 超时意味着 finished_with_result 永不到达，无并发槽竞争，
                    # bool 赋值在 GIL 下原子，后台线程直接置位即可。
                    self._schedule_result_toast_shown = True
                except Exception:
                    pass
                # 让 _schedule_done_signal 携带非零失败数，触发日志面板自动展开
                if not result["failed"]:
                    result["failed"] = 1
            self._post_schedule_log_from_worker(
                f"[定时] 任务「{task.name}」时间点 {slot} 完成："
                f"成功{result['success']}，失败{result['failed']}"
            )
            # 强制兜底恢复标题/任务栏：无论 worker 结束信号是否正确到达，
            # 分段等待循环一定退出（正常结束/超时/排队被停止），这里无条件恢复，
            # 杜绝标题永久停“发送中”
            try:
                self._schedule_done_signal.emit(
                    int(result.get("failed", 0)),
                    int(result.get("success", 0)),
                    list(result.get("failed_list") or []),
                )
            except Exception:
                pass
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

        self.check_update_action = QAction("检查更新", self)
        self.check_update_action.triggered.connect(self._manual_check_update)
        self.tray_menu.addAction(self.check_update_action)
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

    # ------------------------- 在线更新 -------------------------
    def _manual_check_update(self):
        """托盘菜单“检查更新”：手动强制检查，不受间隔/跳过状态限制。"""
        self._start_update_check(manual=True)

    def _on_update_timer_tick(self):
        """长运行期间每小时 tick：距上次成功检查满 24 小时则自动检查。"""
        try:
            if self._update_state.should_periodic_check():
                self._start_update_check(manual=False)
        except Exception:
            logger.exception("定时更新检查 tick 异常")

    def _start_update_check(self, manual: bool):
        # 同一时刻只允许一个检查线程
        if self._update_check_worker is not None \
                and self._update_check_worker.isRunning():
            if manual:
                QMessageBox.information(self, "检查更新", "正在检查更新，请稍候…")
            return
        if manual:
            self.log(f"[更新] 正在检查更新（当前版本 v{APP_VERSION}）…")
            self.check_update_action.setEnabled(False)

        worker = UpdateCheckWorker(
            log_fn=self._post_schedule_log_from_worker, parent=self,
            preferred_mirror=UpdateState.load().preferred_mirror)
        # 信号跨线程自动排队到主线程；槽内只做 GUI/状态操作
        worker.succeeded.connect(
            lambda info: self._on_update_check_result(info, manual))
        worker.failed.connect(
            lambda msg: self._on_update_check_failed(msg, manual))
        worker.finished.connect(self._on_update_check_finished)
        self._update_check_worker = worker
        worker.start()

    def _on_update_check_finished(self):
        """检查线程结束后的统一清理（成功/失败都会到这里）。"""
        try:
            self.check_update_action.setEnabled(True)
        except Exception:
            pass
        worker = self._update_check_worker
        self._update_check_worker = None
        if worker is not None:
            try:
                worker.deleteLater()
            except Exception:
                pass

    def _on_update_check_result(self, info, manual: bool):
        # 成功拿到服务端结果（无论有无新版）都算一次成功检查
        self._update_state.mark_checked()
        if info is None:
            if manual:
                QMessageBox.information(
                    self, "检查更新", f"当前已是最新版本 v{APP_VERSION}。")
            else:
                self.log(f"[更新] 已是最新版本 v{APP_VERSION}")
            return

        if not manual:
            # 自动检查：跳过的版本、已提醒过的版本不再打扰
            if self._update_state.skipped_version == info.version:
                self.log(f"[更新] 新版本 v{info.version} 已被跳过，不再提醒")
                return
            if self._update_state.last_prompted_version == info.version:
                return
            self._update_state.last_prompted_version = info.version
            self._update_state.save()
            self._show_update_toast(info)
        else:
            self.log(f"[更新] 发现新版本 v{info.version}，请查看更新对话框")
            self._show_update_dialog(info)

    def _on_update_check_failed(self, msg: str, manual: bool):
        # 失败不刷新 last_check_at：下个 tick（1 小时后）自动重试
        self.log(f"[更新] 检查更新失败: {msg}")
        if manual:
            QMessageBox.warning(
                self, "检查更新失败",
                f"无法连接更新服务器：\n{msg}\n\n"
                f"请检查网络后重试（将自动尝试 GitHub 直连与国内镜像）。")

    def _show_update_toast(self, info):
        """自动检查发现新版：轻量 Toast，点击后打开更新对话框。"""
        def _on_click():
            try:
                self.show_main_window()
            finally:
                self._show_update_dialog(info)

        ToastNotification.show_toast(
            f"🔄 发现新版本 v{info.version}",
            "点击查看更新内容并立即更新",
            success=True,
            duration=15000,
            parent=self,
            on_click=_on_click,
        )

    def _show_update_dialog(self, info):
        """打开（或前置）更新对话框；非模态，不阻断主界面与日志查看。"""
        existing = self._update_dialog
        if existing is not None and existing.isVisible():
            existing.raise_()
            existing.activateWindow()
            return
        dlg = UpdateDialog(
            info,
            self._update_state,
            is_busy_cb=self._is_other_send_running,
            log_fn=self.log,
            parent=self,
        )
        self._update_dialog = dlg
        dlg.finished.connect(lambda *_: setattr(self, "_update_dialog", None)
                             if self._update_dialog is dlg else None)
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()

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

        # 定时页「立即执行」启动的 worker（ScheduleTab 自己持有）
        schedule_tab = getattr(self, "schedule_tab", None)
        if schedule_tab is not None:
            manual_worker = getattr(schedule_tab, "send_worker", None)
            if manual_worker and isinstance(manual_worker, QThread):
                try:
                    if manual_worker.isRunning():
                        manual_worker.stop()
                except Exception:
                    pass

        dispatcher = getattr(self, "schedule_dispatcher", None)
        if dispatcher is not None:
            try:
                dispatcher.stop()
            except Exception:
                pass

        # 停止监控页轮询 / 手动执行线程
        monitor_tab = getattr(self, "monitor_tab", None)
        if monitor_tab is not None:
            try:
                monitor_tab.shutdown()
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
        manual_worker = None
        schedule_tab = getattr(self, "schedule_tab", None)
        if schedule_tab is not None:
            manual_worker = getattr(schedule_tab, "send_worker", None)
        running = bool(
            (worker and worker.isRunning())
            or (excel_worker and excel_worker.isRunning())
        )
        for w in (schedule_worker, manual_worker):
            if running:
                break
            if w is not None:
                try:
                    if isinstance(w, QThread) and w.isRunning():
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
