from wxauto4 import WeChat
import logging
import os
import tempfile
import threading
import time
import datetime as _dt

from PyQt6.QtGui import QColor, QFont, QFontMetrics, QImage, QPainter

logger = logging.getLogger(__name__)


# pandas 在某些路径下读取 Excel 会得到 Timestamp；导入失败时按 datetime 处理
try:
    import pandas as _pd
    _Timestamp = _pd.Timestamp
except Exception:
    _Timestamp = ()


def format_cell_value(value):
    """规范化单元格值的字符串显示，特别是 datetime.time / datetime.datetime / pandas.Timestamp。
    
    pandas/openpyxl 读取 Excel 时间列时，可能得到 datetime.time 或 Timestamp 对象，
    默认 str() 会显示到秒甚至带毫秒，但用户希望显示到分钟；如果秒非零则保留秒。
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, _dt.time):
        if value.second == 0 and value.microsecond == 0:
            return value.strftime("%H:%M")
        return value.strftime("%H:%M:%S")
    if isinstance(value, _dt.datetime):
        if value.second == 0 and value.microsecond == 0:
            return value.strftime("%Y-%m-%d %H:%M")
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if _Timestamp and isinstance(value, _Timestamp):
        if value.second == 0 and value.microsecond == 0:
            return value.strftime("%Y-%m-%d %H:%M")
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return str(value)


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

    # 全局发送批次锁：手动/定时/监控所有发送批次在此排队串行，
    # 避免多个 worker 同时驱动同一份微信句柄导致窗口争抢发错对象。
    # _BATCH_ACTIVE>0 表示当前有批次持锁（供弹窗 watchdog 门控判断）。
    _BATCH_LOCK = threading.Lock()
    _BATCH_ACTIVE = 0
    _BATCH_STATE_LOCK = threading.Lock()

    @classmethod
    def is_batch_active(cls) -> bool:
        """是否有发送批次正在持锁运行（弹窗 watchdog 据此决定是否扫描）。"""
        with cls._BATCH_STATE_LOCK:
            return cls._BATCH_ACTIVE > 0

    @classmethod
    def acquire_batch(cls, log_fn=None, should_stop=None) -> bool:
        """获取全局发送批次锁（阻塞排队，可被 should_stop 中断）。

        返回 True=已获得锁；False=被 should_stop 中断放弃排队。
        等待期间每 ~5s 提示一次，首次等待即提示"排队等待"。
        """
        waited = False
        last_hint = 0.0
        while True:
            if cls._BATCH_LOCK.acquire(timeout=0.5):
                with cls._BATCH_STATE_LOCK:
                    cls._BATCH_ACTIVE += 1
                return True
            if should_stop and should_stop():
                return False
            now = time.time()
            if not waited:
                waited = True
                last_hint = now
                if log_fn:
                    try:
                        log_fn("⏳ 有其他发送任务正在进行，本任务排队等待其完成后继续…")
                    except Exception:
                        pass
            elif now - last_hint >= 5.0:
                last_hint = now
                if log_fn:
                    try:
                        log_fn("⏳ 仍在排队等待其他发送任务完成…")
                    except Exception:
                        pass

    @classmethod
    def release_batch(cls) -> None:
        """释放全局发送批次锁。未持锁时调用不抛异常。"""
        try:
            with cls._BATCH_STATE_LOCK:
                cls._BATCH_ACTIVE = max(0, cls._BATCH_ACTIVE - 1)
            cls._BATCH_LOCK.release()
        except Exception:
            pass

    def __init__(self):
        self.wx = None
        # 最近一次确认句柄可用的时间戳；超时会先做一次轻量探活
        self._last_healthy_ts = 0.0
        self._own_initialize_called = False
        # ChatInfo 短缓存：避免同一接收人发送时重复 UIA 调用（每次 50-200ms）
        self._chatinfo_cache = None
        self._chatinfo_cache_ts = 0.0
        self._chatinfo_cache_ttl = 1.5  # 秒
        # 发送失败弹窗守护线程：发送期间独立扫描关闭弹窗，
        # 防止模态弹窗阻塞 SendMsg 导致整个批量任务卡死
        self._dismiss_lock = threading.Lock()
        self._watchdog_lock = threading.Lock()
        self._watchdog_thread = None
        self._watchdog_deadline = 0.0
        # 收件人存在性预检查缓存：recipient -> (result, expiry_ts)，
        # 批量发送时同一名字只搜一次，避免重复输入+等待拖慢速度。
        # expiry_ts == 0 表示永久不过期（用户选"永久"档）；否则为到期时间戳。
        # 缓存与 TTL 持久化到 %LOCALAPPDATA%\ExcelSendWx\precheck_cache.json，
        # 重启程序后仍可复用，节省重复搜索耗时。
        self._precheck_cache: dict = {}
        self._precheck_ttl = 300  # 默认有效期 5 分钟（0 = 永久）
        self._precheck_cache_loaded = False
        # 弹窗守护线程挂起计数器：仅收件人预检搜索期间挂起 watchdog
        # （预检要持续读取搜索下拉结果，且该阶段不会出现"发送失败"弹窗）。
        # ChatWith 阶段不挂起——watchdog 通过 _is_suspect_failure_dialog 的
        # ToolSaveBits+Weixin 排除规则识别搜索浮层并跳过，继续防真实失败弹窗
        self._watchdog_suspend = 0

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

    @staticmethod
    def _activate_wechat_window(log_fn=None):
        """在初始化 WeChat 句柄前激活微信主窗口。

        微信窗口最小化到系统托盘时，wxauto4 找不到主窗口句柄，会导致
        WeChatSender.initialize() 失败（定时任务最常见的失败原因）。
        用 Win32 API 把微信主窗口恢复并置前，确保 wxauto4 能正常 attach。

        新版微信 4.0 主窗口类名为 "Qt51514QWindowIcon"，标题为"微信"。
        旧版微信 3.x 主窗口类名为 "WeChatMainWndForPC"。
        """
        try:
            import win32gui
            import win32con
            import ctypes
        except ImportError as exc:
            if log_fn:
                try:
                    log_fn(f"⚠ win32gui 不可用，跳过微信窗口激活: {exc}")
                except Exception:
                    pass
            return False

        candidates = [
            ("Qt51514QWindowIcon", "微信"),      # 微信 4.0 (Weixin.exe)
            ("WeChatMainWndForPC", "微信"),      # 微信 3.x (WeChat.exe)
        ]
        hwnd = 0
        for cls_name, title in candidates:
            try:
                hwnd = win32gui.FindWindow(cls_name, title)
            except Exception:
                hwnd = 0
            if hwnd:
                break

        if not hwnd:
            if log_fn:
                try:
                    log_fn("⚠ 未找到微信主窗口（微信可能未登录或未启动）")
                except Exception:
                    pass
            return False

        try:
            was_visible = bool(win32gui.IsWindowVisible(hwnd))
            was_minimized = False
            if not was_visible:
                # 从系统托盘恢复：SW_RESTORE=9
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                # 微信此前若是最大化，SW_RESTORE 会原样恢复成最大化（看起来像全屏）；
                # 本程序仅因发送需要恢复窗口，归位为普通大小更自然
                try:
                    if win32gui.IsZoomed(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
                except Exception:
                    pass
                was_minimized = True
            # 解除 Windows 前台锁定（SystemParametersInfo, SPI_SETFOREGROUNDLOCKTIMEOUT=0x2001）
            try:
                user32 = ctypes.windll.user32
                user32.SystemParametersInfoW(0x2001, 0, 0, 0)
            except Exception:
                pass
            # 置前；偶尔会因前台锁定抛 OSError，重试一次
            for _ in range(2):
                try:
                    win32gui.SetForegroundWindow(hwnd)
                    break
                except Exception:
                    time.sleep(0.15)
            # 给 wxauto 一点时间稳定 UIA 树
            time.sleep(0.4)
            if log_fn and was_minimized:
                try:
                    log_fn("✅ 已激活微信主窗口（之前最小化到托盘）")
                except Exception:
                    pass
            return True
        except Exception as exc:
            if log_fn:
                try:
                    log_fn(f"⚠ 激活微信窗口失败: {exc}")
                except Exception:
                    pass
            return False

    @staticmethod
    def _ensure_window_on_screen(log_fn=None) -> bool:
        """发送失败重试前的兜底：若微信主窗口大部分在屏幕外/最小化，拉回主屏内。

        窗口位置正常时不做任何操作（保持用户窗口原样，位置不影响发送——
        UIA 定位基于控件实际屏幕坐标，窗口在屏幕哪个位置都能正常发送）。
        但窗口一半以上移出屏幕/最小化时，控件坐标会落在屏幕外或不可见，
        鼠标点击无效，这类失败只有纠正位置才能恢复，故仅在重试路径调用。
        返回 True 表示做了纠正。
        """
        try:
            import win32gui
            import win32con
            import win32api
        except ImportError:
            return False

        candidates = [
            ("Qt51514QWindowIcon", "微信"),      # 微信 4.0 (Weixin.exe)
            ("WeChatMainWndForPC", "微信"),      # 微信 3.x (WeChat.exe)
        ]
        hwnd = 0
        for cls_name, title in candidates:
            try:
                hwnd = win32gui.FindWindow(cls_name, title)
            except Exception:
                hwnd = 0
            if hwnd:
                break
        if not hwnd:
            return False

        try:
            fixed = False
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                # 最小化前若是最大化，SW_RESTORE 会恢复成最大化；此处仅因发送需要
                # 恢复窗口，归位为普通大小，避免微信"全屏"跳动
                try:
                    if win32gui.IsZoomed(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
                except Exception:
                    pass
                fixed = True
                if log_fn:
                    try:
                        log_fn("⚠ 微信窗口处于最小化，已恢复后重试")
                    except Exception:
                        pass

            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            w, h = right - left, bottom - top
            if w <= 0 or h <= 0:
                return fixed

            # 虚拟屏幕（所有显示器整体范围），GetSystemMetrics 索引 76-79
            vx = win32api.GetSystemMetrics(76)
            vy = win32api.GetSystemMetrics(77)
            vw = win32api.GetSystemMetrics(78)
            vh = win32api.GetSystemMetrics(79)

            ix = max(0, min(right, vx + vw) - max(left, vx))
            iy = max(0, min(bottom, vy + vh) - max(top, vy))
            visible_ratio = (ix * iy) / float(w * h)

            if visible_ratio < 0.5:
                # 一半以上在屏幕外：保持原尺寸移到主屏工作区中央
                mw = win32api.GetSystemMetrics(0)
                mh = win32api.GetSystemMetrics(1)
                nx = max(0, (mw - w) // 2)
                ny = max(0, (mh - h) // 2)
                win32gui.MoveWindow(hwnd, nx, ny, w, h, True)
                fixed = True
                if log_fn:
                    try:
                        log_fn(
                            f"⚠ 微信窗口大部分在屏幕外(可见{int(visible_ratio * 100)}%)，"
                            f"已自动移回屏幕中央后重试"
                        )
                    except Exception:
                        pass
            return fixed
        except Exception as exc:
            if log_fn:
                try:
                    log_fn(f"⚠ 检查微信窗口位置失败: {exc}")
                except Exception:
                    pass
            return False

    def initialize(self):
        logger.info("Initializing WeChat client...")
        # 先激活微信主窗口，避免窗口最小化到托盘导致 wxauto 找不到句柄
        # （这是定时任务"初始化微信失败"的主要根因）
        self._activate_wechat_window(self.log)
        try:
            # resize=False：不让 wxauto4 自动把聊天窗口拉大到 800x6000 并移到屏幕左侧。
            # 拉大窗口是为了显示更多消息提高 UIA 识别容错（主要影响读取消息功能），
            # 本程序只用发送链路（搜索+输入框），不依赖大窗口；保持用户窗口原样不跳动。
            # ads=False：关闭 wxauto4 启动时的广告横幅输出。
            self.wx = WeChat(ads=False, resize=False)
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

    def minimize_window(self):
        """整个发送任务结束后最小化微信窗口（隐私保护，任务级一次），失败静默忽略。"""
        try:
            import win32gui
            import win32con
            for cls_name in ("Qt51514QWindowIcon", "WeChatMainWndForPC"):
                try:
                    hwnd = win32gui.FindWindow(cls_name, "微信")
                except Exception:
                    hwnd = 0
                if hwnd:
                    # 若窗口当前是最大化，先归位为普通大小再最小化；否则下次任何恢复
                    # （含 wxauto4 内部 SetActive 的 SW_RESTORE）会回到最大化，看起来像全屏
                    try:
                        if win32gui.IsZoomed(hwnd):
                            win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
                    except Exception:
                        pass
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                    return
        except Exception:
            pass

    @staticmethod
    def _normalize_wechat_window() -> bool:
        """发送前把微信主窗口从「最大化」归位为普通大小。

        wxauto4 每次激活窗口（SetActive/ChatWith 内部）会把最小化的窗口
        SW_RESTORE 恢复；若窗口之前处于最大化状态，恢复后即以最大化（全屏）
        显示，干扰用户。发送是自动化操作，统一在每次发送前把窗口归位为
        普通大小，避免"全屏化"。窗口本就不是最大化时不改动。
        """
        try:
            import win32gui
            import win32con
        except ImportError:
            return False
        for cls_name in ("Qt51514QWindowIcon", "WeChatMainWndForPC"):
            try:
                hwnd = win32gui.FindWindow(cls_name, "微信")
            except Exception:
                hwnd = 0
            if not hwnd:
                continue
            try:
                if win32gui.IsZoomed(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
                    return True
            except Exception:
                pass
            return False
        return False

    def reconnect(self, log_fn=None) -> bool:
        """当检测到微信句柄失效时重新初始化。失败会写日志但不抛异常。"""
        try:
            # 解除旧引用，便于 GC 回收句柄相关资源
            self.wx = None
            self._invalidate_chatinfo_cache()
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
        带 1.5 秒短缓存：同一次发送流程中 ChatWith 后可能连续调用多次 ChatInfo，
        缓存避免冗余 UIA 调用（每次 50-200ms），显著提升发送速度。
        ChatWith 后调用方应通过 _invalidate_chatinfo_cache() 清除缓存。
        """
        if not self.wx:
            return False, None
        # 短缓存命中
        now = time.time()
        if (self._chatinfo_cache is not None
                and now - self._chatinfo_cache_ts < self._chatinfo_cache_ttl):
            return True, self._chatinfo_cache
        try:
            chatinfo = self.wx.ChatInfo()
        except Exception:
            self._chatinfo_cache = None
            return False, None
        if not isinstance(chatinfo, dict):
            self._chatinfo_cache = None
            return False, None
        self._chatinfo_cache = chatinfo
        self._chatinfo_cache_ts = now
        self._last_healthy_ts = now
        return True, chatinfo

    def _invalidate_chatinfo_cache(self):
        """ChatWith 切换了聊天窗口后调用，清除缓存的 ChatInfo。"""
        self._chatinfo_cache = None

    def _press_esc_on_wechat(self):
        """按 Esc 关闭微信搜索面板 / 退出误点进入的页面（视频号、搜一搜等）。

        ChatWith 搜不到联系人时，wxauto4 仍可能点击搜索下拉里的
        "搜索: xxx / 视频号" 等入口，把主界面切走；不清理的话，
        后续重试会在错误页面上连环失败。Esc 将微信恢复到主界面。
        """
        try:
            import win32gui
            import ctypes

            hwnd = 0
            for cls_name in ("Qt51514QWindowIcon", "WeChatMainWndForPC"):
                try:
                    hwnd = win32gui.FindWindow(cls_name, "微信")
                except Exception:
                    hwnd = 0
                if hwnd:
                    break
            if not hwnd:
                return
            # Esc 只作用于前台窗口，先确保微信在前台
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
            time.sleep(0.1)

            VK_ESC = 0x1B
            KEYEVENTF_KEYUP = 0x0002
            INPUT_KEYBOARD = 1

            class KEYBDINPUT(ctypes.Structure):
                _fields_ = [
                    ("wVk", ctypes.c_ushort),
                    ("wScan", ctypes.c_ushort),
                    ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong),
                    ("dwExtraInfo", ctypes.c_void_p),
                ]

            class INPUT(ctypes.Structure):
                class _IU(ctypes.Union):
                    _fields_ = [("ki", KEYBDINPUT)]
                _anonymous_ = ("iu",)
                _fields_ = [("type", ctypes.c_ulong), ("iu", _IU)]

            def _send_key(vk, key_up=False):
                inp = INPUT()
                inp.type = INPUT_KEYBOARD
                inp.ki = KEYBDINPUT(
                    vk, 0, KEYEVENTF_KEYUP if key_up else 0, 0, None
                )
                ctypes.windll.user32.SendInput(
                    1, ctypes.byref(inp), ctypes.sizeof(INPUT)
                )

            _send_key(VK_ESC)
            time.sleep(0.05)
            _send_key(VK_ESC, key_up=True)
            time.sleep(0.15)
            self._invalidate_chatinfo_cache()
        except Exception:
            pass

    def _send_enter_key(self):
        """发送 Enter 键，用于点击微信弹窗的默认'确定'按钮。

        复用 _press_esc_on_wechat 的 SendInput 模式，仅将 VK_ESC 改为 VK_RETURN。
        """
        try:
            import ctypes

            VK_RETURN = 0x0D
            KEYEVENTF_KEYUP = 0x0002
            INPUT_KEYBOARD = 1

            class KEYBDINPUT(ctypes.Structure):
                _fields_ = [
                    ("wVk", ctypes.c_ushort),
                    ("wScan", ctypes.c_ushort),
                    ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong),
                    ("dwExtraInfo", ctypes.c_void_p),
                ]

            class INPUT(ctypes.Structure):
                class _IU(ctypes.Union):
                    _fields_ = [("ki", KEYBDINPUT)]
                _anonymous_ = ("iu",)
                _fields_ = [("type", ctypes.c_ulong), ("iu", _IU)]

            def _send_key(vk, key_up=False):
                inp = INPUT()
                inp.type = INPUT_KEYBOARD
                inp.ki = KEYBDINPUT(
                    vk, 0, KEYEVENTF_KEYUP if key_up else 0, 0, None
                )
                ctypes.windll.user32.SendInput(
                    1, ctypes.byref(inp), ctypes.sizeof(INPUT)
                )

            _send_key(VK_RETURN)
            time.sleep(0.05)
            _send_key(VK_RETURN, key_up=True)
            time.sleep(0.15)
        except Exception:
            pass

    def _find_wechat_dialogs(self):
        """枚举微信进程的所有可见顶层窗口，返回候选弹窗列表。

        返回 (main_hwnd, candidates)，candidates 元素为
        (hwnd, title, class_name, width, height)。main_hwnd 为 0 表示
        微信主窗口未找到（微信未启动），此时 candidates 为空。
        """
        import win32gui
        import win32process

        main_hwnd = 0
        for cls_name in ("Qt51514QWindowIcon", "WeChatMainWndForPC"):
            try:
                main_hwnd = win32gui.FindWindow(cls_name, "微信")
            except Exception:
                main_hwnd = 0
            if main_hwnd:
                break
        if not main_hwnd:
            return 0, []

        try:
            _, wechat_pid = win32process.GetWindowThreadProcessId(main_hwnd)
        except Exception:
            return 0, []

        candidates = []

        def _enum_callback(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if hwnd == main_hwnd:
                return True
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
            except Exception:
                return True
            if pid != wechat_pid:
                return True
            try:
                title = win32gui.GetWindowText(hwnd)
                cls = win32gui.GetClassName(hwnd)
                rect = win32gui.GetWindowRect(hwnd)
                w, h = rect[2] - rect[0], rect[3] - rect[1]
                candidates.append((hwnd, title or "", cls or "", w, h))
            except Exception:
                return True
            return True

        try:
            win32gui.EnumWindows(_enum_callback, None)
        except Exception:
            pass
        return main_hwnd, candidates

    # 发送失败类弹窗标题关键词（微信 3.x 多为 #32770 标准对话框，
    # 4.0 多为 Qt 自绘弹窗，标题可能为空，需结合尺寸兜底判断）
    _DIALOG_TITLE_KEYWORDS = (
        "发送失败", "发送出错", "消息发送失败", "发送中断",
        "操作频繁", "过于频繁", "无法发送", "重新发送", "网络错误",
        "网络异常", "请求超时", "无法连接", "操作太频繁",
        "发送过于频繁", "内容违规",
    )
    _DIALOG_CONFIRM_TEXTS = ("确定", "OK", "好的", "知道了", "重试")

    @classmethod
    def _is_suspect_failure_dialog(cls, title, class_name, w, h, sending=False):
        """判断一个微信顶层窗口是否疑似'发送失败'类弹窗。

        sending=True 表示当前有发送批次持锁（watchdog 门控），此时采用
        更宽松的兜底：微信 4.0 自绘提示框的标题栏常为'微信'等短标题，
        正文才是'发送失败'，仅按关键词/空标题会漏检导致任务卡死。
        """
        if any(k in title for k in cls._DIALOG_TITLE_KEYWORDS):
            return True
        # 排除微信搜索下拉浮层：class 含 ToolSaveBits 且标题为英文 'Weixin'
        # （在搜索框输入收件人名时弹出的搜索结果工具窗）。它不是"发送失败"
        # 弹窗——若被 watchdog 误关会导致真实联系人也被误判为找不到。
        # 真实失败弹窗标题为中文（'微信'/空/'发送失败'）或含失败关键词，
        # 与该特征（英文 Weixin + ToolSaveBits 工具窗）不重叠。
        if "ToolSaveBits" in class_name and title.strip() == "Weixin":
            return False
        # 兜底1：标题为空 + 尺寸较小（排除通话/朋友圈等大窗口），
        # 基本就是模态提示弹窗（微信 4.0 的"发送失败"框无标题栏，
        # 正文直接显示失败文案；类名不限于 Qt，放宽类名避免漏检）
        if not title.strip() and w <= 800 and h <= 600:
            return True
        # 兜底2（发送期间）：任何 Qt 类名的小尺寸顶层窗口，标题为短标题
        # 或含提示类字样，都视为疑似发送失败弹窗（标题栏常显示'微信'，
        # 正文才是失败文案，仅靠关键词会漏检）
        if sending and w <= 800 and h <= 600 and class_name.startswith("Qt"):
            t = title.strip()
            if len(t) <= 8 or any(k in t for k in ("微信", "消息", "提示", "错误", "失败",
                                                   "发送", "注意", "警告", "异常", "重试")):
                return True
        return False

    def _click_dialog_button_uia(self, hwnd) -> bool:
        """通过 UIAutomation 查找弹窗内'知道了/确定'按钮并 Invoke 点击。

        微信 4.0 的"发送失败"框是 Qt 自绘弹窗（无标题栏、无标准 Win32
        Button 子控件），Win32 消息（Enter/Close 等）经常无响应，但 Qt
        控件对 UIA 有暴露。此路径按按钮文本精确定位并触发其 Invoke 模式，
        等价于真实鼠标点击，是关闭该类弹窗最可靠的手段。
        """
        try:
            import comtypes.client
            import comtypes.gen.UIAutomationClient as UIA
        except Exception:
            return False
        try:
            # UIAutomationCore 的 CUIAutomation 类没有注册 ProgID，使用固定 CLSID
            ui_unknown = comtypes.client.CreateObject(
                "{FF48DBA4-60EF-4201-AA87-54103EEF594E}"
            )
            ui = ui_unknown.QueryInterface(UIA.IUIAutomation)
        except Exception:
            return False
        try:
            el = ui.ElementFromHandle(hwnd)
            for text in self._DIALOG_CONFIRM_TEXTS:
                # UIA_NamePropertyId=30005；TreeScope_Descendants=4
                cond = ui.CreatePropertyCondition(30005, text)
                found = el.FindFirst(4, cond)
                if not found:
                    continue
                # UIA_InvokePatternId=10000：找到按钮后触发 Invoke（真实点击）
                pattern = found.GetCurrentPattern(10000)
                if pattern is not None:
                    invoke = pattern.QueryInterface(UIA.IUIAutomationInvokePattern)
                    invoke.Invoke()
                    time.sleep(0.2)
                    return True
            return False
        except Exception:
            return False

    def _click_dialog_bottom_right(self, hwnd) -> bool:
        """终极兜底：模拟鼠标点击弹窗右下角区域。

        Windows 惯例把默认/确认按钮放在右下角，'知道了'按钮通常位于
        弹窗底部偏右。UIA 不可用或未命中按钮文本时，盲点该区域。
        """
        try:
            import win32gui
            import ctypes
            import time as _t

            rect = win32gui.GetWindowRect(hwnd)
            x = rect[0] + int((rect[2] - rect[0]) * 0.85)
            y = rect[1] + int((rect[3] - rect[1]) * 0.82)
            user32 = ctypes.windll.user32
            user32.SetCursorPos(x, y)
            user32.mouse_event(0x0002, 0, 0, 0, 0)  # MOUSEEVENTF_LEFTDOWN
            _t.sleep(0.05)
            user32.mouse_event(0x0004, 0, 0, 0, 0)  # MOUSEEVENTF_LEFTUP
            _t.sleep(0.2)
            return True
        except Exception:
            return False

    def _close_dialog_hwnd(self, hwnd, title, class_name) -> bool:
        """关闭指定弹窗。优先 PostMessage 直接点击'确定'按钮（不依赖前台，
        锁屏/远程桌面断开时也有效），失败再退回置前+SendInput Enter。"""
        import win32gui
        import win32con

        # 1) 找'确定'按钮子窗口（#32770 标准对话框有真实 Button 子控件）
        btn = 0

        def _find_btn(child, _):
            nonlocal btn
            try:
                if win32gui.GetClassName(child) == "Button":
                    txt = win32gui.GetWindowText(child)
                    if any(t in txt for t in self._DIALOG_CONFIRM_TEXTS):
                        btn = child
                        return False
            except Exception:
                pass
            return True

        try:
            win32gui.EnumChildWindows(hwnd, _find_btn, None)
        except Exception:
            pass
        if btn:
            try:
                win32gui.PostMessage(btn, win32con.BM_CLICK, 0, 0)
                time.sleep(0.15)
                if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                    return True
            except Exception:
                pass

        # 2) 对弹窗本身发 Enter 键消息（Qt 自绘弹窗无 Button 子控件，
        #    但默认按钮通常持有焦点，Enter 可触发）
        try:
            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
            time.sleep(0.15)
            if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                return True
        except Exception:
            pass

        # 2.1) UIAutomation 精确定位'知道了/确定'按钮并 Invoke（微信 4.0
        #      Qt 自绘弹窗最可靠关闭方式，等价真实鼠标点击按钮）
        try:
            if self._click_dialog_button_uia(hwnd):
                if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                    return True
        except Exception:
            pass

        # 2.5) 直接请求关闭窗口 / 发送 IDOK 命令 / Esc 键
        # （Qt 自绘弹窗对标准 Button 与 Enter 可能无响应，这三条消息覆盖面更广）
        try:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            time.sleep(0.15)
            if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                return True
        except Exception:
            pass
        try:
            win32gui.PostMessage(hwnd, win32con.WM_COMMAND, 1, 0)  # IDOK=1
            time.sleep(0.15)
            if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                return True
        except Exception:
            pass
        try:
            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)
            time.sleep(0.15)
            if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                return True
        except Exception:
            pass

        # 3) 退回置前后发 SendInput Enter（需要交互桌面，锁屏时可能无效）
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass
        time.sleep(0.1)
        self._send_enter_key()
        time.sleep(0.15)
        try:
            if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                return True
        except Exception:
            pass

        # 4) 终极兜底：模拟鼠标点击弹窗右下角（'知道了'按钮惯例位置）
        try:
            if self._click_dialog_bottom_right(hwnd):
                time.sleep(0.2)
        except Exception:
            pass
        try:
            return not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd)
        except Exception:
            return True

    def _dismiss_send_failure_dialog(self) -> bool:
        """检测并关闭微信'发送失败'类弹窗。返回是否检测到弹窗。

        线程安全（_dismiss_lock 防重入），供发送线程与守护线程并发调用。
        发现的候选窗口均记入日志（标题/类名/尺寸），便于事后定位弹窗特征。
        """
        try:
            import win32gui  # noqa: F401
            import win32process  # noqa: F401
        except ImportError:
            return False

        with self._dismiss_lock:
            try:
                _, candidates = self._find_wechat_dialogs()
            except Exception:
                return False
            if not candidates:
                return False

            sending = self.is_batch_active()
            found = False
            missed = []
            for hwnd, title, cls, w, h in candidates:
                if not self._is_suspect_failure_dialog(title, cls, w, h, sending=sending):
                    missed.append((title, cls, w, h))
                    continue
                found = True
                try:
                    closed = self._close_dialog_hwnd(hwnd, title, cls)
                    self.log(
                        f"⚠️ 检测到微信弹窗(title={title!r}, class={cls!r}, "
                        f"{w}x{h})，自动关闭{'成功' if closed else '可能失败'}"
                    )
                except Exception as exc:
                    self.log(f"⚠️ 关闭微信弹窗异常: {exc}")
            # 诊断：发送期间存在未命中的候选窗口时记录其特征（节流 10s），
            # 便于发现新弹窗标题变体导致的漏检，避免任务静默卡死无从定位
            if sending and missed:
                now = time.time()
                if now - getattr(self, "_last_missed_diag_ts", 0.0) >= 10.0:
                    self._last_missed_diag_ts = now
                    self.log(
                        f"🔎 发送期间未命中的微信窗口(非弹窗或未知弹窗): "
                        f"{missed}"
                    )
            return found

    def _ensure_dialog_watchdog(self):
        """确保弹窗守护线程在运行（常驻 daemon，幂等启动）。

        每 0.5 秒扫描一次微信弹窗；仅当有发送批次持锁
        （is_batch_active()）时才实际执行关闭，避免误关用户手动
        在微信里操作产生的弹窗。线程随进程退出，无需 stop。
        """
        with self._watchdog_lock:
            t = self._watchdog_thread
            if t is not None and t.is_alive():
                return
            self._watchdog_thread = threading.Thread(
                target=self._dialog_watchdog_loop,
                daemon=True,
                name="wx-dialog-watchdog",
            )
            self._watchdog_thread.start()

    def _dialog_watchdog_loop(self):
        while True:
            try:
                # 发送批次持锁时才扫描关闭（门控），空闲/用户手动操作不干预；
                # 收件人预检期间挂起（计数器>0）；ChatWith 的搜索浮层则靠
                # ToolSaveBits+Weixin 排除规则识别跳过，不在此挂起
                if self.is_batch_active() and self._watchdog_suspend == 0:
                    self._dismiss_send_failure_dialog()
            except Exception:
                pass
            time.sleep(0.5)

    def _send_msg_with_dialog_retry(self, content, recipient, stop_event=None) -> bool:
        """发送消息，检测发送失败弹窗并重试。最多 3 次。

        弹窗失败通常因频率限制/网络问题，重试间隔 1 秒；
        SendMsg 返回失败间隔 0.5 秒。
        """
        max_retries = 3
        for attempt in range(max_retries):
            if self._stop_requested(stop_event):
                return False
            resp = self.wx.SendMsg(content)
            # 等待弹窗可能出现
            if self._wait_or_stopped(0.3, stop_event):
                return False
            if self._dismiss_send_failure_dialog():
                if attempt < max_retries - 1:
                    self.log(f"⚠️ {recipient} 发送失败(微信弹窗)，第{attempt + 1}/{max_retries}次重试...")
                    if self._wait_or_stopped(1.0, stop_event):
                        return False
                    continue
                self.log(f"❌ {recipient} 弹窗重试{max_retries}次仍失败")
                return False
            if resp is not None and not resp:
                if attempt < max_retries - 1:
                    self.log(f"⚠️ {recipient} SendMsg返回失败({resp})，第{attempt + 1}/{max_retries}次重试...")
                    if self._wait_or_stopped(0.5, stop_event):
                        return False
                    continue
                self.log(f"❌ {recipient} SendMsg重试{max_retries}次仍失败")
                return False
            return True
        return False

    def _send_file_with_dialog_retry(self, file_path, recipient, stop_event=None) -> bool:
        """发送文件，检测发送失败弹窗并重试。最多 3 次。"""
        max_retries = 3
        for attempt in range(max_retries):
            if self._stop_requested(stop_event):
                return False
            resp = self.wx.SendFiles(file_path)
            if self._wait_or_stopped(0.3, stop_event):
                return False
            if self._dismiss_send_failure_dialog():
                if attempt < max_retries - 1:
                    self.log(f"⚠️ {recipient} 文件发送失败(微信弹窗)，第{attempt + 1}/{max_retries}次重试...")
                    if self._wait_or_stopped(1.0, stop_event):
                        return False
                    continue
                self.log(f"❌ {recipient} 文件弹窗重试{max_retries}次仍失败")
                return False
            if resp is not None and not resp:
                if attempt < max_retries - 1:
                    self.log(f"⚠️ {recipient} SendFiles返回失败({resp})，第{attempt + 1}/{max_retries}次重试...")
                    if self._wait_or_stopped(0.5, stop_event):
                        return False
                    continue
                self.log(f"❌ {recipient} SendFiles重试{max_retries}次仍失败")
                return False
            return True
        return False

    def _precheck_contact_exists(self, recipient: str) -> bool:
        """发送前预验证收件人在微信里能搜到联系人项。

        原因：wxauto4 的 ChatWith 搜不到联系人时，会误点搜索下拉里的
        「搜索网络结果」入口，把微信主界面切到网页搜索页（Esc 无效），
        导致后续重试连环失败。本方法改用 UIA 直接往微信搜索框输入名字、
        读取搜索结果下拉：
        - 存在匹配的联系人/群/公众号项（mmui::SearchContentCellView）→ 放行
        - 不存在 → 按 Esc 关闭搜索浮层并返回 False，调用方跳过该收件人，
          绝不调用 ChatWith，从根上避免弹出搜索网页。
        任何异常/找不到搜索框时返回 True 放行，由原逻辑兜底。

        预检期间会挂起弹窗守护线程（self._watchdog_suspend），避免把微信搜索
        下拉浮层（Qt 工具窗，标题 'Weixin'）误判为"发送失败"弹窗关掉，
        否则搜索结果一出现就被关，真实联系人也会被误判为不存在。
        """
        key = str(recipient or "").strip()
        if not key:
            return True
        self._load_precheck_cache()
        now = time.time()
        cached = self._precheck_cache.get(key)
        if cached and (cached[1] == 0 or cached[1] > now):
            return cached[0]

        # 预检期间挂起弹窗守护线程（finally 中先 Esc 关浮层再恢复计数）
        self._watchdog_suspend += 1
        result = True
        try:
            import comtypes.client
            import comtypes.gen.UIAutomationClient as UIA
            import win32gui

            hwnd = 0
            for cls_name in ("Qt51514QWindowIcon", "WeChatMainWndForPC"):
                try:
                    hwnd = win32gui.FindWindow(cls_name, "微信")
                except Exception:
                    hwnd = 0
                if hwnd:
                    break
            if not hwnd:
                self.log("⚠ 找不到微信窗口，跳过收件人预检查")
                return True

            # 仅当微信窗口不可见/最小化到托盘时才恢复并置前：同一批连续预检
            # 通常窗口已在前台，若每次都激活会反复抢焦点又拖慢速度
            need_activate = True
            try:
                # 最小化窗口 IsWindowVisible 仍为 True，需同时排除 IsIconic
                need_activate = not (
                    win32gui.IsWindowVisible(hwnd) and not win32gui.IsIconic(hwnd)
                )
            except Exception:
                pass
            if need_activate:
                # 隐藏窗口 UIA 读不到内容，先恢复并置前
                try:
                    self._activate_wechat_window(self.log)
                except Exception:
                    pass

            # 定位微信顶部搜索框（name='搜索' 或 class 含 XValidatorTextEdit）。
            # 窗口刚从托盘恢复时 UIA 树可能仍在懒加载，搜索框未立即暴露，
            # 需置前+短等重试；窗口已在前台时无需置前，改用更短的等待。
            # 仍找不到则按 Ctrl+F 强制聚焦搜索框兜底。
            import ctypes
            edit = None
            ui = None
            root = None
            # 已在前台时缩短等待：UIA 树已就绪，搜索框立即可读
            wait_sleep = 0.3 if need_activate else 0.05
            for _try in range(3):
                if need_activate:
                    try:
                        win32gui.SetForegroundWindow(hwnd)
                    except Exception:
                        pass
                time.sleep(wait_sleep)
                ui_unknown = comtypes.client.CreateObject(
                    "{FF48DBA4-60EF-4201-AA87-54103EEF594E}"
                )
                ui = ui_unknown.QueryInterface(UIA.IUIAutomation)
                root = ui.ElementFromHandle(hwnd)
                cond = ui.CreatePropertyCondition(30005, "搜索")
                edit = root.FindFirst(4, cond)
                if edit:
                    break
                # 按名字找不到：遍历 class 含 XValidatorTextEdit 的元素
                all_cond = ui.CreateTrueCondition()
                found_all = root.FindAll(4, all_cond)
                for i in range(found_all.Length):
                    el = found_all.GetElement(i)
                    try:
                        if "XValidatorTextEdit" in str(el.CurrentClassName or ""):
                            edit = el
                            break
                    except Exception:
                        continue
                if edit:
                    break
                # 兜底：Ctrl+F 聚焦搜索框后下一轮重试
                try:
                    ctypes.windll.user32.keybd_event(0x11, 0, 0, 0)
                    ctypes.windll.user32.keybd_event(0x46, 0, 0, 0)
                    ctypes.windll.user32.keybd_event(0x46, 0, 2, 0)
                    ctypes.windll.user32.keybd_event(0x11, 0, 2, 0)
                except Exception:
                    pass
                time.sleep(0.5)
            if not edit:
                self.log("⚠ 找不到微信搜索框，跳过收件人预检查")
                return True

            try:
                vp = edit.GetCurrentPattern(10002)  # ValuePatternId
                vp.QueryInterface(UIA.IUIAutomationValuePattern).SetValue(key)
            except Exception:
                self.log("⚠ 收件人预检查输入失败，放行")
                return True

            # 轮询读取搜索下拉：微信渲染慢/卡顿时多等几次再判定，
            # 避免只读一次因渲染延迟误判"找不到"而跳过真实联系人。
            # 结果通常渲染在主窗口树；个别情况下搜索浮层是独立 Qt 工具窗
            # （class 常为 Qt51514QWindowToolSaveBits，标题 'Weixin'），
            # 主窗口树未命中时才枚举顶层工具窗补扫，避免每轮 FindWindow 开销。
            kl = key.lower()
            result = False
            deadline = time.time() + 1.2
            while time.time() < deadline:
                all_cond = ui.CreateTrueCondition()
                # 1) 先扫主窗口树（最常见情况）
                if self._scan_tree_for_contact(root, all_cond, kl):
                    result = True
                    break
                # 2) 主窗口树未命中：补扫顶层 Qt 工具窗（搜索浮层独立成窗时）
                try:
                    _, tool_windows = self._find_wechat_dialogs()
                except Exception:
                    tool_windows = []
                for _thw, _tt, _tc, _tw, _th in tool_windows:
                    try:
                        tree = ui.ElementFromHandle(_thw)
                    except Exception:
                        continue
                    if self._scan_tree_for_contact(tree, all_cond, kl):
                        result = True
                        break
                if result:
                    break
                time.sleep(0.2)
        except Exception as e:
            self.log(f"⚠ 收件人预检查异常，放行: {e}")
            result = True
        finally:
            # 先关掉搜索浮层（Esc），再恢复守护线程，避免浮层残留被误关
            try:
                self._press_esc_on_wechat()
            except Exception:
                pass
            self._watchdog_suspend = max(0, self._watchdog_suspend - 1)
            # 统一写缓存（含放行/失败路径），有效期内同收件人不再重复搜索。
            # ttl <= 0 表示永久不过期，用户手动清理；否则存到期时间戳
            ttl = self._precheck_ttl
            expiry = 0 if ttl <= 0 else time.time() + ttl
            self._precheck_cache[key] = (result, expiry)
            self._persist_precheck_cache()
        return result

    @staticmethod
    def _precheck_cache_path() -> str:
        """预检缓存持久化文件：%LOCALAPPDATA%\\ExcelSendWx\\precheck_cache.json"""
        import os
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        d = os.path.join(base, "ExcelSendWx")
        try:
            os.makedirs(d, exist_ok=True)
        except Exception:
            pass
        return os.path.join(d, "precheck_cache.json")

    def _load_precheck_cache(self):
        """从磁盘加载预检缓存与 TTL，剔除已过期条目。幂等（只读一次）。"""
        if self._precheck_cache_loaded:
            return
        self._precheck_cache_loaded = True
        try:
            import json
            path = self._precheck_cache_path()
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            ttl = data.get("ttl_seconds")
            if isinstance(ttl, (int, float)) and ttl >= 0:
                self._precheck_ttl = int(ttl)
            now = time.time()
            cleaned = {}
            for k, v in (data.get("entries") or {}).items():
                try:
                    if not isinstance(v, (list, tuple)) or len(v) < 2:
                        continue
                    result, expiry = bool(v[0]), float(v[1])
                except Exception:
                    continue
                if expiry == 0 or expiry > now:
                    cleaned[k] = (result, expiry)
            self._precheck_cache = cleaned
            if cleaned:
                self.log(
                    f"📂 已加载收件人预检缓存 {len(cleaned)} 条"
                    f"（有效期={'永久' if self._precheck_ttl <= 0 else str(self._precheck_ttl) + ' 秒'}）"
                )
        except Exception:
            self._precheck_cache = {}

    def _persist_precheck_cache(self):
        """把缓存与 TTL 原子写回磁盘（临时文件 + replace）。"""
        try:
            import json
            import os
            path = self._precheck_cache_path()
            tmp = path + ".tmp"
            data = {
                "ttl_seconds": self._precheck_ttl,
                "entries": {k: list(v) for k, v in self._precheck_cache.items()},
            }
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
            os.replace(tmp, path)
        except Exception:
            pass

    def set_precheck_ttl(self, seconds: int) -> None:
        """设置预检缓存有效期（秒）；0 = 永久不过期，需手动清理。立即持久化。"""
        self._load_precheck_cache()
        self._precheck_ttl = max(0, int(seconds or 0))
        self._persist_precheck_cache()

    def clear_precheck_cache(self) -> int:
        """清空预检缓存并持久化，返回清除条数。"""
        self._load_precheck_cache()
        n = len(self._precheck_cache)
        self._precheck_cache = {}
        self._persist_precheck_cache()
        return n

    def get_precheck_cache_stats(self):
        """返回 (缓存条数, ttl_seconds)，供 GUI 展示。"""
        self._load_precheck_cache()
        return len(self._precheck_cache), self._precheck_ttl

    @staticmethod
    def _scan_tree_for_contact(tree, all_cond, kl: str) -> bool:
        """扫描一个 UIA 树，返回是否命中收件人搜索结果项。

        WeChat 的联系人/群/公众号搜索结果项均为 mmui::SearchContentCellView，
        且元素名包含收件人关键字。任何异常视为未命中。
        """
        if tree is None:
            return False
        try:
            items = tree.FindAll(4, all_cond)
        except Exception:
            return False
        for i in range(items.Length):
            el = items.GetElement(i)
            try:
                cls = str(el.CurrentClassName or "")
                name = str(el.CurrentName or "").strip()
            except Exception:
                continue
            if cls == "mmui::SearchContentCellView" and kl and kl in name.lower():
                return True
        return False

    def _safe_chatwith(self, recipient: str, exact: bool = False) -> bool:
        if not self.wx:
            return False
        # ChatWith 会再次打开搜索下拉；watchdog 通过 _is_suspect_failure_dialog
        # 的 ToolSaveBits+Weixin 排除规则识别该浮层并跳过，无需挂起
        try:
            self.wx.ChatWith(recipient, exact=exact)
            # 切换了聊天窗口，清除 ChatInfo 缓存
            self._invalidate_chatinfo_cache()
            return True
        except Exception:
            # 搜不到目标时 ChatWith 内部可能已把界面切到搜索结果/视频号页，
            # 按 Esc 恢复主界面，避免后续重试在错误页面上连环失败
            self._press_esc_on_wechat()
            return False

    def open_chat(self, recipient: str, chat_delay: float = 0.3,
                  stop_event=None) -> bool:
        """切换到收件人聊天窗口并等待就绪（含重试）。

        供多内容发送时调用一次：之后的 text/image/attachment/custom 步骤
        直接用 fast_mode 在当前窗口发送，不再重复搜索-切换-确认。
        """
        if not self.wx:
            if not self.initialize():
                return False
        if not self._precheck_contact_exists(recipient):
            self.log(f"❌ 微信中找不到联系人「{recipient}」，跳过发送（避免误点搜索网页）")
            return False
        max_retries = 3
        for attempt in range(max_retries):
            if self._stop_requested(stop_event):
                return False
            switched = self._safe_chatwith(recipient, exact=False)
            if not switched:
                if attempt < max_retries - 1:
                    self._ensure_window_on_screen(self.log)
                    if self._wait_or_stopped(0.2, stop_event):
                        return False
                continue
            if self._wait_or_stopped(chat_delay, stop_event):
                return False
            ok, chatinfo = self._safe_chatinfo()
            if not ok:
                if self.reconnect(self.log):
                    if self._safe_chatwith(recipient, exact=False):
                        self._wait_or_stopped(chat_delay, stop_event)
                        ok, chatinfo = self._safe_chatinfo()
            current_chat = chatinfo.get('chat_name', '') if chatinfo else ''
            if self._is_target_chat(current_chat, recipient):
                return True
            if attempt < max_retries - 1:
                self._press_esc_on_wechat()
                self._ensure_window_on_screen(self.log)
                if self._wait_or_stopped(0.2, stop_event):
                    return False
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

        # 启动弹窗守护线程：发送期间独立扫描关闭'发送失败'模态弹窗，
        # 防止其阻塞 SendMsg 的 UIA 调用导致批量任务整体卡死
        self._ensure_dialog_watchdog()
        # 进入发送前，先清理可能残留的发送失败弹窗（上一次发送可能遗留）
        self._dismiss_send_failure_dialog()
        # 预验证收件人存在：微信搜不到时跳过，避免 ChatWith 误点搜索网页
        if not self._precheck_contact_exists(recipient):
            self.log(f"❌ 微信中找不到联系人「{recipient}」，跳过发送（避免误点搜索网页）")
            return False
        # 发送前把微信窗口从最大化归位普通大小，避免 wxauto4 激活后"全屏化"
        self._normalize_wechat_window()

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
                        if not self._send_msg_with_dialog_retry(content, recipient, stop_event):
                            return False
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
                            # 兜底：窗口若移出屏幕/最小化，ChatWith 点击会无效
                            self._ensure_window_on_screen(self.log)
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
                        if not self._send_msg_with_dialog_retry(content, recipient, stop_event):
                            return False
                        self.log(f"✅ 消息发送成功")
                        return True
                    else:
                        if attempt < max_retries - 1:
                            self.log(f"窗口切换失败，按 Esc 恢复主界面后重新搜索 ({attempt+1}/{max_retries})")
                            # ChatWith 可能点到了搜索下拉（视频号/搜一搜入口），
                            # 先恢复主界面再重试，避免在错误页面上连环失败
                            self._press_esc_on_wechat()
                            # 兜底：窗口若移出屏幕/最小化，重试仍会点击无效
                            self._ensure_window_on_screen(self.log)
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
        # 第二次整体重试前：最后兜底检查窗口位置（移出屏幕/最小化则纠正）
        self._ensure_window_on_screen(self.log)
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

        self.log(f"发送文件: {file_path}")

        # 启动弹窗守护线程（同 send_message，防止模态弹窗阻塞卡死）
        self._ensure_dialog_watchdog()
        # 进入发送前，先清理可能残留的发送失败弹窗（上一次发送可能遗留）
        self._dismiss_send_failure_dialog()
        # 预验证收件人存在：微信搜不到时跳过，避免 ChatWith 误点搜索网页
        if not self._precheck_contact_exists(recipient):
            self.log(f"❌ 微信中找不到联系人「{recipient}」，跳过发送（避免误点搜索网页）")
            return False
        # 发送前把微信窗口从最大化归位普通大小，避免 wxauto4 激活后"全屏化"
        self._normalize_wechat_window()

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
                        if not self._send_file_with_dialog_retry(file_path, recipient, stop_event):
                            return False
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
                            # 兜底：窗口若移出屏幕/最小化，ChatWith 点击会无效
                            self._ensure_window_on_screen(self.log)
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
                        if not self._send_file_with_dialog_retry(file_path, recipient, stop_event):
                            return False
                        self.log(f"✅ 图片发送成功")
                        return True

                    if attempt < max_retries - 1:
                        self.log(f"窗口切换失败，按 Esc 恢复主界面后重新搜索 ({attempt + 1}/{max_retries})")
                        # ChatWith 可能点到了搜索下拉（视频号/搜一搜入口），
                        # 先恢复主界面再重试
                        self._press_esc_on_wechat()
                        # 兜底：窗口若移出屏幕/最小化，重试仍会点击无效
                        self._ensure_window_on_screen(self.log)
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
            "" if value is None else format_cell_value(value)
            for value in table_data.get("headers", [])
        ]
        rows = table_data.get("rows", [])
        if not headers:
            raise ValueError("图片数据缺少表头")

        column_count = len(headers)
        normalized_rows = []
        for row in rows:
            values = [
                "" if value is None else format_cell_value(value)
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
        stop_event=None,
        fast_mode=False
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
                    # fast_mode=True 时所有图片都直接在当前窗口发（窗口已预切换）
                    fast_mode=fast_mode or index > 0,
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
        stop_event=None,
        fast_mode=False
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
                first_send=(i == 0 and not fast_mode),
                chat_delay=chat_delay,
                fast_mode=fast_mode,
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
