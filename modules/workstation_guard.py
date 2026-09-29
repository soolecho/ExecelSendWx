r"""
电脑锁定守护（定时发送场景）。

背景：定时任务触发时，Windows 会话必须处于"解锁"状态，否则 UIA 无法操作
微信窗口，发送必然失败。Windows 安全桌面（Winlogon）限制导致普通程序无法
对已锁定的会话自动解锁（向锁屏界面注入密码不可行也不安全），因此采用等效方案：

1. 守护线程运行期间持续阻止"自动锁定"：
   - SetThreadExecutionState 防止系统空闲睡眠；
   - 关闭屏幕保护程序（屏保 + "恢复时显示登录屏幕"是夜间自动锁定的最常见来源）；
   - 尝试清零组策略 InactivityTimeoutSecs（无操作自动锁屏，需管理员权限）。
2. 定时任务发送完成后：若当天近期（默认 15 分钟内）没有其他待触发任务，
   自动调用 LockWorkStation() 重新锁定电脑。

限制：任务触发时若电脑已被锁定（例如用户手动 Win+L），无法自动解锁，
只能由 ScheduleSendWorker 输出警告日志。

所有日志通过注入的 log_fn 输出（MainWindow 会用信号投递到主线程），
本模块不直接操作任何 QWidget。
"""
import ctypes
import json
import logging
import os
import tempfile
import threading
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Win32 DLL 句柄：模块级预绑定，避免在子线程里用 ctypes.windll 动态查找。
#   ctypes.windll 是惰性加载且会缓存失败状态，某些 Python 版本 / 沙箱里
#   在守护线程内首次查函数会报 "function 'XXX' not found"，进程级预绑定最稳。
#   另外：SetThreadExecutionState 属于 kernel32.dll，**不是** user32.dll，
#   之前错写为 windll.user32.SetThreadExecutionState，偶发找不到就崩。
# ---------------------------------------------------------------------------
try:
    _user32 = ctypes.WinDLL("user32", use_last_error=True)
    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
except Exception as _exc:
    logger.error("ctypes WinDLL 加载失败, 守护功能不可用: %s", _exc)
    _user32 = None
    _kernel32 = None


# SetThreadExecutionState 标志
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

# SystemParametersInfo 调用标志
_SPIF_UPDATEINIFILE = 0x01
_SPIF_SENDCHANGE = 0x02

SPI_GETSCREENSAVEACTIVE = 0x000E
SPI_SETSCREENSAVEACTIVE = 0x000F
SPI_GETSCREENSAVESECURE = 0x00F6
SPI_SETSCREENSAVESECURE = 0x00F7

# 任务完成后多久内还有任务就不锁屏（秒）
_RELOCK_WINDOW_SECONDS = 15 * 60

_GPO_KEY_PATH = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
_GPO_VALUE_NAME = "InactivityTimeoutSecs"


def _call(fn, *args, default=None):
    """以 try/except 调用 DLL 函数，失败返回 default，不抛异常。"""
    if fn is None:
        return default
    try:
        return fn(*args)
    except Exception as exc:
        logger.warning("Win32 API 调用失败 %s: %s", getattr(fn, "__name__", fn), exc)
        return default


def _base_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        base = Path(local_app_data) / "ExcelSendWx"
    else:
        base = Path(tempfile.gettempdir()) / "ExcelSendWx"
    base.mkdir(parents=True, exist_ok=True)
    return base


def _settings_path() -> Path:
    return _base_dir() / "guard.json"


def load_guard_enabled() -> bool:
    try:
        with _settings_path().open("r", encoding="utf-8") as fp:
            data = json.load(fp)
        return bool(data.get("keep_unlocked", False))
    except (OSError, json.JSONDecodeError, ValueError):
        return False


def save_guard_enabled(enabled: bool) -> None:
    try:
        with _settings_path().open("w", encoding="utf-8") as fp:
            json.dump(
                {"keep_unlocked": bool(enabled)},
                fp,
                ensure_ascii=False,
                indent=2,
            )
    except OSError as exc:
        logger.warning("保存守护开关失败: %s", exc)


def is_workstation_locked() -> bool:
    """探测当前输入桌面是否为 Winlogon（会话已锁定）。

    锁定时输入桌面是 Winlogon，普通进程没有权限打开它；
    未锁定时 OpenInputDesktop 成功返回桌面句柄。
    """
    if _user32 is None:
        return False
    try:
        # DESKTOP_SWITCHDESKTOP = 0x0100
        hdesk = _call(_user32.OpenInputDesktop, 0, False, 0x0100, default=None)
        if not hdesk:
            return True
        _call(_user32.CloseDesktop, hdesk)
        return False
    except Exception:
        # 探测失败按未锁定处理，避免误报
        return False


def lock_workstation() -> bool:
    """锁定工作站，等价于按下 Win+L。"""
    if _user32 is None:
        return False
    try:
        return bool(_call(_user32.LockWorkStation, default=0))
    except Exception as exc:
        logger.warning("LockWorkStation 失败: %s", exc)
        return False


def _read_inactivity_timeout():
    """读取组策略"无操作自动锁定"秒数；未配置返回 None。"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, _GPO_KEY_PATH, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, _GPO_VALUE_NAME)
            return int(value)
    except OSError:
        return None


def _write_inactivity_timeout(seconds) -> bool:
    """写入组策略锁屏秒数（需管理员权限）。seconds=None 表示删除该值。"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, _GPO_KEY_PATH, 0, winreg.KEY_SET_VALUE) as key:
            if seconds is None:
                try:
                    winreg.DeleteValue(key, _GPO_VALUE_NAME)
                except FileNotFoundError:
                    return True
            else:
                winreg.SetValueEx(key, _GPO_VALUE_NAME, 0, winreg.REG_DWORD, int(seconds))
        return True
    except OSError as exc:
        logger.info("无法修改 InactivityTimeoutSecs（需要管理员权限）: %s", exc)
        return False


def _spi_get_bool(spi_get):
    if _user32 is None:
        return None
    try:
        value = ctypes.c_int(0)
        ok = _call(
            _user32.SystemParametersInfoW,
            spi_get, 0, ctypes.byref(value), 0,
            default=0,
        )
        return bool(value.value) if ok else None
    except Exception:
        return None


def _spi_set_bool(spi_set, value) -> bool:
    if _user32 is None:
        return False
    try:
        ok = _call(
            _user32.SystemParametersInfoW,
            spi_set, int(bool(value)), None,
            _SPIF_UPDATEINIFILE | _SPIF_SENDCHANGE,
            default=0,
        )
        return bool(ok)
    except Exception:
        return False


class WorkstationGuard:
    """防锁定守护线程。与 GUI 完全解耦，日志只通过注入的 log_fn 输出。"""

    def __init__(self, log_fn=None, poll_interval=30.0):
        self._log_fn = log_fn
        self._poll_interval = max(5.0, float(poll_interval))
        self._stop_event = threading.Event()
        self._thread = None
        self._saved = None  # (screensave_active, inactivity_orig)
        self._gpo_fix_applied = False
        # 防重复锁定：一次守护周期内只锁一次，用户解锁后不再自动锁定
        self._relock_done = False
        self._relock_lock = threading.Lock()

    # ------------------------------------------------------------------
    def _log(self, message: str) -> None:
        logger.info(message)
        if self._log_fn:
            try:
                self._log_fn(message)
            except Exception:
                pass

    def is_running(self) -> bool:
        t = self._thread
        return bool(t and t.is_alive())

    def start(self) -> None:
        if self.is_running():
            return
        self._stop_event.clear()
        self._relock_done = False  # 新守护周期重置锁定标志
        self._thread = threading.Thread(
            target=self._run, name="WorkstationGuard", daemon=True
        )
        self._thread.start()
        self._log("[电脑守护] 已开启：阻止电脑自动锁定/休眠，任务完成后自动锁定")

    def stop(self) -> None:
        if not self.is_running():
            return
        self._stop_event.set()
        t = self._thread
        if t and t.is_alive():
            t.join(timeout=self._poll_interval)
        self._log("[电脑守护] 已关闭，电脑锁定策略已恢复原状")

    # ------------------------------------------------------------------
    def _snapshot_and_disable_lock_sources(self) -> None:
        # 首次：读取并保存原始状态；后续只重申设置，避免每 30 秒重复读注册表
        if self._saved is None:
            active = _spi_get_bool(SPI_GETSCREENSAVEACTIVE)
            inactivity = _read_inactivity_timeout()
            self._saved = (active, inactivity)
            logger.info(
                "WorkstationGuard 初始状态: 屏保=%s, InactivityTimeoutSecs=%s",
                active, inactivity,
            )
            # 首次清零无操作锁屏策略（尽力而为，失败不影响其余手段）
            if not self._gpo_fix_applied and inactivity:
                if _write_inactivity_timeout(0):
                    self._gpo_fix_applied = True
                    self._log(
                        f"[电脑守护] 已临时清零无操作锁屏策略（原值 {inactivity} 秒）"
                    )
        # 关闭屏幕保护程序（连带其"恢复时锁定"失效）——每次都重申，防止被系统/组策略改回
        _spi_set_bool(SPI_SETSCREENSAVEACTIVE, False)

    def _restore_lock_sources(self) -> None:
        saved = self._saved
        if saved is None:
            return
        active, inactivity = saved
        if active is not None:
            _spi_set_bool(SPI_SETSCREENSAVEACTIVE, active)
        if self._gpo_fix_applied:
            if _write_inactivity_timeout(inactivity if inactivity is not None else None):
                self._gpo_fix_applied = False
        self._saved = None
        logger.info("WorkstationGuard 已恢复: 屏保=%s, InactivityTimeoutSecs=%s", active, inactivity)

    def _run(self) -> None:
        try:
            self._snapshot_and_disable_lock_sources()
            while not self._stop_event.wait(self._poll_interval):
                # SetThreadExecutionState 是线程级状态，由守护线程持续刷新。
                # 注意：该 API 在 kernel32.dll（之前误写 user32 导致偶发 not found）。
                # ES_DISPLAY_REQUIRED 防止屏幕熄灭/屏保启动（仅 ES_SYSTEM_REQUIRED
                # 只防睡眠，不防屏保——这是"操作中也进屏保"的根因）。
                _call(
                    _kernel32.SetThreadExecutionState if _kernel32 else None,
                    ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED,
                )
                # 屏保/策略可能被用户或其他程序改回，周期性重申
                self._snapshot_and_disable_lock_sources()
        finally:
            # ES_CONTINUOUS 不带其它标志 = 清除本线程的阻止状态
            _call(
                _kernel32.SetThreadExecutionState if _kernel32 else None,
                ES_CONTINUOUS,
            )
            self._restore_lock_sources()

    # ------------------------------------------------------------------
    def maybe_relock_after_task(self, store, now: datetime = None) -> bool:
        """定时任务完成后调用：若近期无其他待触发任务则锁定电脑。

        返回 True 表示本次执行了锁定。
        可在任意线程调用（LockWorkStation 线程安全；日志经 log_fn 投递）。

        防重复：一次守护周期内只锁一次。用户解锁后不会再被自动锁定，
        避免出现"解锁→立刻又被锁定"的死循环。
        """
        # 已锁定过一次，不再重复锁定（防止用户解锁后又立刻被锁）
        if self._relock_done:
            return False

        with self._relock_lock:
            # double-check：拿到锁后再确认一次
            if self._relock_done:
                return False

            now = now or datetime.now()
            upcoming = []
            try:
                for task in store.load_all():
                    if not task.enabled:
                        continue
                    if task.repeat_mode == "weekly":
                        if now.isoweekday() not in set(task.days):
                            continue
                    elif task.repeat_mode == "once":
                        # 一次性任务只有执行日期当天才会触发
                        if now.strftime("%Y-%m-%d") not in set(getattr(task, "run_dates", None) or []):
                            continue
                    for slot in task.times:
                        try:
                            hh, mm = slot.split(":", 1)
                            slot_dt = now.replace(
                                hour=int(hh), minute=int(mm), second=0, microsecond=0
                            )
                        except (ValueError, AttributeError):
                            continue
                        delta = (slot_dt - now).total_seconds()
                        if 0 <= delta <= _RELOCK_WINDOW_SECONDS:
                            upcoming.append(slot)
            except Exception as exc:
                self._log(f"[电脑守护] 检查后续任务失败: {exc}")
                return False

            if upcoming:
                self._log(
                    f"[电脑守护] 近期还有任务（{'、'.join(sorted(set(upcoming)))}），暂不锁定"
                )
                return False

            if is_workstation_locked():
                return False

            if lock_workstation():
                self._relock_done = True  # 标记已锁定，本周期不再重复
                self._log("[电脑守护] 发送完成，已自动锁定电脑（本周期仅此一次，解锁后不再自动锁定）")
                return True
            return False
