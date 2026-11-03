"""uia_bridge —— WeChatUIA（wechatauto 1.2.5.1）桥接层。

将微信 4.x UIA 驱动封装为 WeChatSender 需要的窄接口，进程内单例 + 线程锁串行。
替换 wxauto4 依赖后，wechat_sender 只 import 本模块，不再直接依赖 wxauto4。

关键设计
--------
1. 线程安全：uiautomation 控件对象**不能跨线程传递**（COM 单元模型）。每个公开
   方法进入时先 CoInitializeEx（``auto.InitializeUIAutomationInCurrentThread``，
   幂等），并避免复用上个线程缓存的主窗口控件：所有读操作经 ``_find_main()`` 在
   当前线程重锚；写操作（open_chat/send_text/send_files）内部本来就先
   ``ensure_window()`` 刷新 ``_win``。
2. 文件/图片发送采用「两步粘贴 + 一次 Enter」（v1.5.0 批处理合并约束）：
   ① CF_HDROP 写入剪贴板 → 输入框 Ctrl+V 出文件卡片/图片草稿
   ② CF_UNICODETEXT 写入剪贴板 → Ctrl+V 文字追加在卡片后
   ③ 一次 Enter 全部发出。
   （微信输入框对"同一剪贴板同时含 CF_HDROP+CF_UNICODETEXT"只渲染文件、丢弃
   文字，必须分两步粘贴。）
3. 剪贴板写入带 OpenClipboard 多轮重试：剪贴板被输入法/Office 等进程持有时
   等对方释放，写后读回自校验。
"""

from __future__ import annotations

import ctypes
import difflib
import logging
import os
import threading
import time
from typing import List, Optional

import uiautomation as auto
from wechatauto.param import WxParam
from wechatauto.uia_driver import (
    SESSION_LIST_AIDS,
    WeChatUIA,
    _aid_hit,
    _clean_chat_name,
    _find_by,
)

logger = logging.getLogger(__name__)

CF_HDROP = 15
CF_UNICODETEXT = 13
# 全局内存标志：GMEM_MOVEABLE | GMEM_ZEROINIT
_GMEM_MOVEABLE_ZEROINIT = 0x0042

# ---------------------------------------------------------------------------
# @ 成员（群聊）：键入 "@昵称" 后微信弹出 MentionPopover，点选候选项落成 chip。
# 纯 UIA 实现（与 wechatauto/ui/chatbox.py 的 AtMenu 同一套锚点），不依赖 OCR。
# ---------------------------------------------------------------------------
AT_POPOVER_CLASS = "mmui::XPopover"
AT_POPOVER_NAME = "Weixin"
AT_POPOVER_AID = "MentionPopover"
# @所有人 在不同版本/语言包里的候选名（依次尝试）
AT_ALL_ALIASES = ("所有人", "全部成员", "All", "Everyone")


def _mention_key(name: str) -> str:
    """@ 匹配用的归一化名字：去空格（微信弹层里昵称不含空格）。"""
    return "".join(str(name or "").split())


def _sendkeys_literal(text: str) -> str:
    """把文本转成 uiautomation SendKeys 字面量（花括号是唯一需要转义的字符）。"""
    return "".join("{{}" if ch == "{" else "{}}" if ch == "}" else ch
                   for ch in str(text or ""))


# ---------------------------------------------------------------------------
# 剪贴板（模块级，供两步粘贴使用；带 OpenClipboard 重试防剪贴板争用）
# ---------------------------------------------------------------------------
def _open_clipboard_with_retry(max_wait: float = 1.5) -> bool:
    """OpenClipboard 多轮重试：剪贴板被其他进程持有时等对方释放再写。"""
    user32 = ctypes.windll.user32
    try:
        user32.OpenClipboard.argtypes = [ctypes.c_void_p]
        user32.OpenClipboard.restype = ctypes.wintypes.BOOL
    except Exception:
        pass
    deadline = time.time() + max_wait
    while True:
        try:
            if user32.OpenClipboard(None):
                return True
        except Exception:
            pass
        if time.time() >= deadline:
            return False
        time.sleep(0.03)


def clipboard_set_files(paths) -> bool:
    """把本地文件以 CF_HDROP 格式写入剪贴板（多文件）。返回是否成功。

    与 wechatauto.guia.copy_files_to_clipboard 同技术（纯 ctypes，无 pywin32），
    另加 OpenClipboard 重试。仅统计真实存在的文件，空列表返回 False。
    """
    file_paths = [os.path.abspath(p) for p in paths if os.path.isfile(p)]
    if not file_paths:
        return False

    class DROPFILES(ctypes.Structure):
        _fields_ = [
            ("pFiles", ctypes.c_uint),
            ("pt", ctypes.wintypes.POINT),
            ("fNC", ctypes.c_int),
            ("fWide", ctypes.c_int),
        ]

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    try:
        user32.EmptyClipboard.restype = ctypes.wintypes.BOOL
        user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
        user32.SetClipboardData.restype = ctypes.c_void_p
        kernel32.GlobalAlloc.restype = ctypes.c_void_p
        kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    except Exception:
        pass
    try:
        if not _open_clipboard_with_retry():
            logger.warning("clipboard_set_files: OpenClipboard 重试超时（剪贴板被占用）")
            return False
        user32.EmptyClipboard()
        df = DROPFILES()
        df.pFiles = ctypes.sizeof(DROPFILES)
        df.fWide = 1
        raw = (ctypes.string_at(ctypes.byref(df), ctypes.sizeof(DROPFILES))
               + ("\0".join(file_paths) + "\0").encode("utf-16-le") + b"\0\0")
        h = kernel32.GlobalAlloc(_GMEM_MOVEABLE_ZEROINIT, len(raw))
        if not h:
            user32.CloseClipboard()
            return False
        dst = kernel32.GlobalLock(h)
        if not dst:
            user32.CloseClipboard()
            return False
        try:
            ctypes.memmove(dst, raw, len(raw))
        finally:
            kernel32.GlobalUnlock(h)
        user32.SetClipboardData(CF_HDROP, h)
        user32.CloseClipboard()
        return True
    except Exception as e:
        logger.warning("clipboard_set_files 失败: %s", e)
        try:
            user32.CloseClipboard()
        except Exception:
            pass
        return False


def clipboard_set_text(text: str, max_wait: float = 1.5) -> bool:
    """把文本以 CF_UNICODETEXT 写入剪贴板，写后读回自校验。返回是否成功。"""
    text = text or ""
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    try:
        user32.EmptyClipboard.restype = ctypes.wintypes.BOOL
        user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
        user32.SetClipboardData.restype = ctypes.c_void_p
        user32.GetClipboardData.argtypes = [ctypes.c_uint]
        user32.GetClipboardData.restype = ctypes.c_void_p
        user32.IsClipboardFormatAvailable.argtypes = [ctypes.c_uint]
        kernel32.GlobalAlloc.restype = ctypes.c_void_p
        kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    except Exception:
        pass
    try:
        if not _open_clipboard_with_retry(max_wait):
            logger.warning("clipboard_set_text: OpenClipboard 重试超时（剪贴板被占用）")
            return False
        user32.EmptyClipboard()
        data = (text + "\0").encode("utf-16-le")
        h = kernel32.GlobalAlloc(_GMEM_MOVEABLE_ZEROINIT, len(data))
        if not h:
            user32.CloseClipboard()
            return False
        dst = kernel32.GlobalLock(h)
        if not dst:
            user32.CloseClipboard()
            return False
        try:
            ctypes.memmove(dst, data, len(data))
        finally:
            kernel32.GlobalUnlock(h)
        if not user32.SetClipboardData(CF_UNICODETEXT, h):
            user32.CloseClipboard()
            return False
        # 读回自校验（仍在 OpenClipboard 会话内）：防写入被静默吞掉
        try:
            if user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
                h2 = user32.GetClipboardData(CF_UNICODETEXT)
                if h2:
                    p2 = kernel32.GlobalLock(h2)
                    if p2:
                        try:
                            if ctypes.wstring_at(p2) != text:
                                logger.warning("clipboard_set_text 读回内容不一致")
                        finally:
                            kernel32.GlobalUnlock(h2)
        except Exception:
            pass
        user32.CloseClipboard()
        return True
    except Exception as e:
        logger.warning("clipboard_set_text 失败: %s", e)
        try:
            user32.CloseClipboard()
        except Exception:
            pass
        return False


# ---------------------------------------------------------------------------
# UiaBridge
# ---------------------------------------------------------------------------
class UiaBridge:
    """WeChatUIA 封装单例（进程内共享），线程锁串行所有驱动操作。

    线程安全要点：WeChatUIA 控件对象不能跨线程传递，每个公开方法进入时先
    CoInitializeEx（幂等），读操作经 ``_find_main()`` 在当前线程重锚主窗口，
    写操作交给驱动（内部先 ensure_window 刷新 ``_win``）。
    """

    _instance = None
    _instance_lock = threading.Lock()

    def __init__(self, timeout: float = 15.0, search_timeout: float = 2.0):
        self._lock = threading.RLock()
        self._timeout = timeout
        self._search_timeout = search_timeout
        self._driver = None

    @classmethod
    def instance(cls) -> "UiaBridge":
        with cls._instance_lock:
            if cls._instance is None:
                cls._instance = cls()
        return cls._instance

    # ---------------- 内部 ----------------
    @staticmethod
    def _com_init() -> None:
        """当前线程 COM 初始化（幂等；不同线程必须各自初始化才能用 UIA）。"""
        try:
            auto.InitializeUIAutomationInCurrentThread()
        except Exception:
            pass

    def _get_driver(self) -> WeChatUIA:
        self._com_init()
        if self._driver is None:
            self._driver = WeChatUIA(timeout=self._timeout,
                                     search_timeout=self._search_timeout)
        return self._driver

    # ---------------- 公开 API ----------------
    def ensure(self) -> bool:
        """幂等热激活：确保微信主窗可访问并置前（含 Qt accessibility gate 热写）。

        窗口过小自动放大等批次级保障仍由 WeChatSender 负责（原 _ensure_wechat_window_size
        逻辑不动），这里只管"窗口可用且在前台"。
        """
        with self._lock:
            try:
                return self._get_driver().ensure_window(wake=True)
            except Exception as e:
                logger.warning("uia_bridge.ensure 失败: %s", e)
                return False

    def current_chat(self) -> Optional[str]:
        """当前打开会话名（输入框 Name 剥离占位尾巴）；未打开会话/读不到返回 None。

        只读探测：在当前线程重锚主窗口后读输入框，不激活窗口、不抢焦点。
        """
        with self._lock:
            try:
                driver = self._get_driver()
                win = driver._find_main()
                if win is None:
                    return None
                e = driver._chat_input(win)
                if e is None:
                    return None
                return _clean_chat_name(getattr(e, "Name", "") or "")
            except Exception as e:
                logger.warning("uia_bridge.current_chat 失败: %s", e)
                return None

    def open_chat(self, recipient: str, retries: int = 2) -> bool:
        """搜索并打开联系人/群聊，成功后校验输入框 Name。失败 = 收件人不存在/不可达。

        驱动内部 ``_collect_results`` 已过滤「搜索网络结果」，不会误点搜索网页；
        keyword 支持 wxid→昵称/备注映射。此调用天然充当"联系人存在性预检"。
        """
        with self._lock:
            try:
                return self._get_driver().open_chat(
                    str(recipient or "").strip(), retries=max(1, retries))
            except Exception as e:
                logger.warning("uia_bridge.open_chat(%s) 异常: %s", recipient, e)
                return False

    # ---------------- @ 成员（群聊） ----------------
    @staticmethod
    def _at_popover(driver, timeout: float = 1.2):
        """等待并返回 @ 候选弹层（MentionPopover）；没有（如私聊）返回 None。"""
        win = None
        try:
            win = driver._find_main()
        except Exception:
            win = None
        if win is None:
            win = getattr(driver, "_win", None)
        if win is None:
            return None
        try:
            popover = win.WindowControl(
                ClassName=AT_POPOVER_CLASS,
                Name=AT_POPOVER_NAME,
                AutomationId=AT_POPOVER_AID,
            )
            if popover.Exists(timeout, 0.15):
                return popover
        except Exception:
            pass
        return None

    @staticmethod
    def _popover_items(popover):
        """枚举弹层里的候选项（ListControl 子项，兜底用直接子项）。"""
        items = []
        try:
            lst = popover.ListControl()
            if lst is not None:
                items = list(lst.GetChildren())
        except Exception:
            items = []
        if not items:
            try:
                items = list(popover.GetChildren())
            except Exception:
                items = []
        return [c for c in items if _mention_key(getattr(c, "Name", ""))]

    @staticmethod
    def _click_at_item(item) -> bool:
        """点击候选项（Click 失败时退回 Invoke）。"""
        try:
            item.Click()
            return True
        except Exception:
            pass
        try:
            item.GetInvokePattern().Invoke()
            return True
        except Exception:
            return False

    def _pick_at_item(self, popover, name: str, exact_only: bool = False) -> bool:
        """在弹层里选中目标成员：唯一候选直接选，否则按名字（含别名）匹配。

        ``exact_only=True`` 时不做"包含"匹配：输入名字前缀阶段（弹层里还有
        多个候选）用，避免把"张三丰"当成"张三"点错人。
        """
        items = self._popover_items(popover)
        if not items:
            return False
        key = _mention_key(name)
        aliases = [key] + [_mention_key(a) for a in AT_ALL_ALIASES]
        if len(items) == 1:
            # 唯一候选：只有名字确有关联时才点，避免点错人
            only = _mention_key(getattr(items[0], "Name", ""))
            at_all_keys = {_mention_key(a) for a in AT_ALL_ALIASES}
            if (key and (key in only or only in key)) or (
                    key in at_all_keys and only in at_all_keys):
                return self._click_at_item(items[0])
            return False
        for want in aliases:
            for item in items:
                if _mention_key(getattr(item, "Name", "")) == want:
                    if self._click_at_item(item):
                        return True
        if exact_only:
            return False
        # 候选项可能带备注/后缀（如 "张三（备注）"）：退化为包含匹配
        for want in aliases:
            for item in items:
                nm = _mention_key(getattr(item, "Name", ""))
                if want and want in nm:
                    if self._click_at_item(item):
                        return True
        return False

    @staticmethod
    def _has_at_targets(at_members, at_all) -> bool:
        """是否配置了 @ 目标（成员或 @所有人）。"""
        raw = at_members
        if isinstance(raw, str):
            raw = [raw]
        if any(str(m or "").strip() for m in (raw or [])):
            return True
        return bool(at_all)

    @staticmethod
    def _main_hwnd(driver) -> int:
        """当前微信主窗口句柄（取不到返回 0）。"""
        try:
            win = driver._find_main()
        except Exception:
            win = None
        if win is None:
            win = getattr(driver, "_win", None)
        try:
            return int(getattr(win, "NativeWindowHandle", 0) or 0)
        except Exception:
            return 0

    @staticmethod
    def _foreground_hwnd() -> int:
        """系统当前前台窗口句柄（失败返回 0）。"""
        try:
            return int(ctypes.windll.user32.GetForegroundWindow())
        except Exception:
            return 0

    def _focus_chat_input(self, driver, e) -> bool:
        """让微信主窗处于前台并把键盘焦点交给聊天输入框，返回是否 SetFocus 成功。

        不能依赖 ``_click_ctrl`` 的屏幕坐标点击：输入框被置顶窗口（视频会议、
        图片/视频预览等）遮挡时，点击会落到遮挡窗口上，后续 SendKeys 的
        "@昵称" 也送给那个窗口——表现就是输入框里没有 @、候选弹层不出现、
        清理时又像"把内容删掉了"。UIA SetFocus 直接作用于输入框控件，
        不经过屏幕坐标。
        """
        hwnd = self._main_hwnd(driver)
        if hwnd:
            try:
                driver._force_foreground(hwnd)
            except Exception:
                pass
            time.sleep(0.2)
        focused = False
        for _ in range(2):
            try:
                e.SetFocus()
                focused = True
            except Exception:
                pass
            time.sleep(0.2)
            if not hwnd or self._foreground_hwnd() == hwnd:
                return focused
            try:
                driver._force_foreground(hwnd)
            except Exception:
                pass
            time.sleep(0.15)
        logger.warning(
            "uia_bridge: 微信主窗口未处于前台（可能被其它置顶窗口遮挡），@ 可能输入失败"
        )
        return focused

    @staticmethod
    def _input_value(e) -> Optional[str]:
        """读输入框当前文本；读不到返回 None。"""
        try:
            return WeChatUIA._read_edit_value(e)
        except Exception:
            return None

    @staticmethod
    def _clear_chat_input(e) -> None:
        """清空输入框（此时还没有 mention chip，Ctrl+A 安全）。"""
        try:
            e.SetFocus()
        except Exception:
            pass
        try:
            e.SendKeys("{Ctrl}a{Delete}", waitTime=0.05)
        except Exception:
            pass
        time.sleep(0.1)

    def _undo_typed(self, driver, e, count: int) -> None:
        """@ 插入失败时清理已键入的 "@昵称"（重新聚焦，先关弹层再退格）。"""
        try:
            self._focus_chat_input(driver, e)
        except Exception:
            pass
        try:
            e.SendKeys("{Esc}", waitTime=0.05)
        except Exception:
            pass
        for _ in range(max(0, int(count))):
            try:
                e.SendKeys("{Back}", waitTime=0.02)
            except Exception:
                return

    @staticmethod
    def _typed_prefix(name: str) -> str:
        """在 @ 弹层里用于过滤候选的键入字符：只保留字母/数字/汉字。

        微信候选过滤对标点很敏感：候选名是「Lin威廉  .」时键入完整原名
        「@Lin威廉.」，过滤结果为空、弹层直接关闭——既选不到人，看起来
        又像"输入了 @ 但没选人名"。只键入有效字符（如「Lin威廉」）时弹层
        会稳定停在唯一候选上，再按去空格后的全名精确匹配即可点选。
        """
        return "".join(ch for ch in str(name or "") if ch.isalnum())

    def _insert_one_mention(self, driver, e, name: str):
        """插入单个 @，返回 (status, typed_len)。

        status: True=已落成 chip；False=弹层里没有该成员；None=没有出现弹层。
        typed_len: 当前输入框里由本函数键入的字符数（含 "@"），供失败清理退格。
        """
        typed = self._typed_prefix(name) or _mention_key(name)
        pending = 1  # "@" 本身
        self._focus_chat_input(driver, e)
        try:
            e.SendKeys("@", waitTime=0.05)
        except Exception:
            return None, 0
        popover = self._at_popover(driver, timeout=2.0)
        if popover is None:
            return None, pending
        if self._pick_at_item(popover, name, exact_only=True):
            return True, pending
        for ch in typed:
            try:
                e.SendKeys(_sendkeys_literal(ch), waitTime=0.05)
            except Exception:
                return None, pending
            pending += 1
            time.sleep(0.2)
            popover = self._at_popover(driver, timeout=1.2)
            if popover is None:
                # 过滤过头（该字符把候选全部滤掉、弹层关闭）：退格回到有候选的状态
                for _ in range(2):
                    try:
                        e.SendKeys("{Back}", waitTime=0.05)
                    except Exception:
                        break
                    pending = max(1, pending - 1)
                    time.sleep(0.25)
                    popover = self._at_popover(driver, timeout=1.2)
                    if popover is not None:
                        break
                break
            if self._pick_at_item(popover, name, exact_only=True):
                return True, pending
        if popover is None:
            return None, pending
        if self._pick_at_item(popover, name):
            return True, pending
        return False, pending

    def _insert_mentions(self, driver, e, at_members=None, at_all: bool = False) -> int:
        """把 @成员 / @所有人 逐个落成输入框 chip，返回成功插入个数。

        每个名字按「键入 @ → 等弹层 → 渐进键入名字 → 匹配点选」执行：
        - 弹层出现且匹配 → 落成 chip；
        - 弹层没出现但输入框已有文本（私聊/版本不支持 @）→ 清理并整体降级为
          普通发送，绝不把字面量 "@昵称" 发出去；
        - 弹层没出现且输入框仍为空（多为焦点被遮挡窗口抢走）→ 重新聚焦后最多
          重试 3 次，避免静默丢掉 @。
        """
        names = []
        seen = set()
        raw_members = at_members
        if isinstance(raw_members, str):
            raw_members = [raw_members]
        for raw in list(raw_members or []):
            nm = str(raw or "").strip()
            if not nm:
                continue
            if _mention_key(nm) not in seen:
                seen.add(_mention_key(nm))
                names.append(nm)
        if at_all:
            alias = AT_ALL_ALIASES[0]
            if _mention_key(alias) not in seen:
                seen.add(_mention_key(alias))
                names.append(alias)

        inserted = 0
        for nm in names:
            status = None
            pending = 0
            for attempt in range(3):
                status, pending = self._insert_one_mention(driver, e, nm)
                if status is True:
                    inserted += 1
                    time.sleep(0.25)  # 等 chip 落位、焦点回到输入框
                    break
                if status is None:
                    value = self._input_value(e)
                    # 已插入的 mention chip 读回是 U+FFFC，去掉后为空说明
                    # 本次 "@名字" 根本没落进输入框（焦点被抢），要重试；
                    # 去掉 chip 后仍有文本才是私聊/不支持 @ 的降级场景。
                    clean = (value or "").replace("\ufffc", "").strip()
                    if clean:
                        # 私聊/不支持 @：弹层不会出现，清掉已键入文本后整体降级
                        logger.info(
                            "uia_bridge: 未出现 @ 候选弹层，跳过 @%s（按普通消息发送）", nm)
                        self._undo_typed(driver, e, pending)
                        return inserted
                    # 输入框仍是空的：按键没落进输入框（焦点被抢），重新聚焦再试
                    if attempt < 2:
                        logger.debug("uia_bridge: @%s 未落进输入框，重新聚焦重试", nm)
                        self._clear_chat_input(e)
                        continue
                    logger.info(
                        "uia_bridge: @%s 输入未能落进输入框（焦点被抢占），跳过", nm)
                    self._clear_chat_input(e)
                    break
                # status=False：弹层在，但候选里没有该成员
                logger.info("uia_bridge: @ 候选里没有「%s」，已清理", nm)
                self._undo_typed(driver, e, pending)
                break
            if status is not True:
                continue
        return inserted

    def send_text(self, text: str, at_members=None, at_all: bool = False) -> bool:
        """在已打开会话的输入框发送文本（粘贴 + 回读相似度校验 + 回车）。

        带 ``at_members`` / ``at_all`` 时先把 @ chip 插进输入框再粘文字；
        不带 @ 时走驱动原路径（行为完全不变）。
        """
        has_at = self._has_at_targets(at_members, at_all)
        with self._lock:
            try:
                if not has_at:
                    return self._get_driver().send_text(text or "")
                if not self.ensure():
                    return False
                driver = self._get_driver()
                e = driver._chat_input()
                if e is None:
                    logger.warning("uia_bridge.send_text: 找不到聊天输入框")
                    return False
                if not self._focus_chat_input(driver, e):
                    logger.warning("uia_bridge.send_text: 聚焦输入框失败")
                    return False
                self._clear_chat_input(e)
                inserted = self._insert_mentions(driver, e, at_members, at_all)
                text = text or ""
                if text:
                    if not clipboard_set_text(text):
                        logger.warning("uia_bridge.send_text: 写文字剪贴板失败")
                        return False
                    try:
                        e.SendKeys("{Ctrl}v", waitTime=0.05)
                    except Exception as exc:
                        logger.warning("uia_bridge.send_text: 粘贴文字异常: %s", exc)
                        return False
                    time.sleep(0.4)
                    if not self._text_landed(e, text):
                        logger.warning("uia_bridge.send_text: 文字未落进输入框，清空不回车")
                        try:
                            e.SendKeys("{Ctrl}a{Delete}", waitTime=0.05)
                        except Exception:
                            pass
                        return False
                elif inserted == 0:
                    return False
                try:
                    e.SendKeys("{Enter}", waitTime=0.05)
                except Exception as exc:
                    logger.warning("uia_bridge.send_text: 回车发送异常: %s", exc)
                    return False
                time.sleep(0.2)
                return True
            except Exception as e:
                logger.warning("uia_bridge.send_text 异常: %s", e)
                return False

    def send_files(self, paths, text: Optional[str] = None,
                   verify_text: bool = True, text_first: bool = False,
                   at_members=None, at_all: bool = False) -> bool:
        """两步粘贴 + 一次 Enter 发送文件/图片，可混合文字（批处理合并约束）。

        两步粘贴的先后决定合并消息里的排列顺序（微信输入框按粘贴顺序排列，
        对"同一剪贴板同时含 CF_HDROP+CF_UNICODETEXT"只渲染文件丢弃文字，必须分两步粘）：
        - text_first=False（默认）：① CF_HDROP 文件 → 卡片 → ② CF_UNICODETEXT 文字追加在卡片后
        - text_first=True ：① CF_UNICODETEXT 文字 → ② CF_HDROP 文件卡片插在文字后
        - 最后一次 Enter 全部发出。

        输入框用 UIA ``_chat_input()`` 锚定（不用 OCR/坐标），点击用
        ``_click_ctrl`` 拿焦点（不吃 WS_EX_TRANSPARENT 的亏）。文字粘贴后可
        回读校验（读不到值按放行），校验失败则清空不回车。文字在前时在校验时
        输入框只有文字，读回最可靠。

        带 ``at_members`` / ``at_all`` 时，@ chip 插在清空输入框之后、粘贴
        文字/文件之前；带 chip 后不再用 Ctrl+A 清空（会删掉 mention chip）。
        """
        with self._lock:
            if not self.ensure():
                return False
            driver = self._get_driver()
            e = driver._chat_input()
            if e is None:
                logger.warning("uia_bridge.send_files: 找不到聊天输入框")
                return False
            file_paths = [os.path.abspath(p) for p in paths if os.path.isfile(p)]
            if not file_paths:
                logger.warning("uia_bridge.send_files: 没有有效文件路径")
                return False
            has_at = self._has_at_targets(at_members, at_all)
            if not has_at:
                if not driver._click_ctrl(e):
                    logger.warning("uia_bridge.send_files: 点击输入框拿焦点失败")
                    return False
            # 清空输入框残留草稿，保证从干净状态开始
            time.sleep(0.15)
            self._clear_chat_input(e)
            # @ 成员/所有人：必须在清空之后、粘贴之前插入（chip 不能被 Ctrl+A 删）
            if has_at:
                # 有 @ 时不能用坐标点击拿焦点（输入框被置顶窗口遮挡会点到别处，
                # 后续 "@昵称" 也会丢给那个窗口），改用 UIA SetFocus。
                if not self._focus_chat_input(driver, e):
                    logger.warning("uia_bridge.send_files: 聚焦输入框失败")
                    return False
            self._insert_mentions(driver, e, at_members, at_all)
            text = (text or "").strip()

            # 文字在前：先粘文字（此刻输入框只有文字，回读校验最可靠），再粘文件
            if text_first and text:
                if not clipboard_set_text(text):
                    logger.warning("uia_bridge.send_files: 写文字剪贴板失败")
                    return False
                try:
                    e.SendKeys("{Ctrl}v", waitTime=0.05)
                except Exception as exc:
                    logger.warning("uia_bridge.send_files: 粘贴文字异常: %s", exc)
                    return False
                time.sleep(0.4)
                if verify_text and not self._text_landed(e, text):
                    logger.warning("uia_bridge.send_files: 文字未落进输入框，清空不回车")
                    try:
                        e.SendKeys("{Ctrl}a{Delete}", waitTime=0.05)
                    except Exception:
                        pass
                    return False

            # 文件/图片（一次 CF_HDROP 粘贴，卡片顺序 = file_paths 顺序）
            if not clipboard_set_files(file_paths):
                logger.warning("uia_bridge.send_files: 写文件剪贴板失败")
                return False
            try:
                e.SendKeys("{Ctrl}v", waitTime=0.05)
            except Exception as exc:
                logger.warning("uia_bridge.send_files: 粘贴文件异常: %s", exc)
                return False
            time.sleep(1.0)  # 等文件卡片/图片缩略图渲染

            # 文件在前（默认）：文字追加在卡片后
            if not text_first and text:
                if not clipboard_set_text(text):
                    logger.warning("uia_bridge.send_files: 写文字剪贴板失败")
                    return False
                try:
                    e.SendKeys("{Ctrl}v", waitTime=0.05)
                except Exception as exc:
                    logger.warning("uia_bridge.send_files: 粘贴文字异常: %s", exc)
                    return False
                time.sleep(0.4)
                if verify_text and not self._text_landed(e, text):
                    logger.warning("uia_bridge.send_files: 文字未落进输入框，清空不回车")
                    try:
                        e.SendKeys("{Ctrl}a{Delete}", waitTime=0.05)
                    except Exception:
                        pass
                    return False
            # 一次 Enter 全部发出
            try:
                e.SendKeys("{Enter}", waitTime=0.05)
            except Exception as exc:
                logger.warning("uia_bridge.send_files: 回车发送异常: %s", exc)
                return False
            time.sleep(1.2)  # 等微信异步上传/发送
            return True

    @staticmethod
    def _text_landed(e, text: str) -> bool:
        """回读输入框确认文字已落进（与文件卡片共存时读回可能不含卡片，只核对文字）。

        读不到值按放行——没有探测能力不等于内容错了（对齐 uia_driver._paste_landed）。
        """
        try:
            got = WeChatUIA._read_edit_value(e)
        except Exception:
            return True
        if got is None:
            return True
        want = WeChatUIA._compare_norm(text)
        # @ chip 在输入框里读回可能是 U+FFFC 或 "@昵称" 前缀，去掉后再比对
        have = WeChatUIA._compare_norm(got).replace("\ufffc", "")
        if not want:
            return True
        if want in have or have.endswith(want):
            return True
        try:
            score = difflib.SequenceMatcher(None, want, have, autojunk=False).ratio()
        except Exception:
            score = 0.0
        need = float(getattr(WxParam, "SEND_CONTENT_RATIO", 0.6) or 0)
        if need <= 0:
            return True
        return score >= need

    def get_chat_list(self, limit: int = 200) -> List[str]:
        """尽力枚举会话列表名称（UIA session_list 子树）；失败返回 []。"""
        with self._lock:
            if not self.ensure():
                return []
            driver = self._get_driver()
            names: List[str] = []
            try:
                node = _find_by(
                    driver._win,
                    lambda c: _aid_hit(getattr(c, "AutomationId", ""),
                                       SESSION_LIST_AIDS),
                    max_depth=30)
                if node is None:
                    return names
                for child in node.GetChildren():
                    nm = (getattr(child, "Name", "") or "").strip()
                    if nm:
                        names.append(nm)
                        if len(names) >= limit:
                            break
            except Exception as e:
                logger.warning("uia_bridge.get_chat_list 失败: %s", e)
            return names

    def is_online(self) -> bool:
        """微信进程在线且主窗可物化（只读探测，不激活、不抢焦点）。"""
        with self._lock:
            self._com_init()
            try:
                if self._driver is not None:
                    if self._driver._find_main() is not None:
                        return True
                return WeChatUIA.is_running()
            except Exception:
                return False
