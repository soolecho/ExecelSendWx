"""Windows 任务栏进度条封装（ITaskbarList3）。

在任务栏程序图标上显示发送进度（绿色进度条/红色错误/黄色暂停），
即使主窗口被微信窗口挡住也能直观看到发送进度。

纯 ctypes 实现，无需 comtypes/pywin32 额外依赖；非 Windows 或初始化
失败时所有方法静默降级（no-op），不影响主程序。
"""
from __future__ import annotations

import ctypes
import logging
import sys
from ctypes import wintypes

logger = logging.getLogger(__name__)

# 进度状态
TBPF_NOPROGRESS = 0
TBPF_INDETERMINATE = 1
TBPF_NORMAL = 2
TBPF_ERROR = 3
TBPF_PAUSED = 4

_CLSID_TaskbarList = "{56FDF344-FD6D-11d0-958A-006097C9A090}"
_IID_ITaskbarList3 = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"

_CLSCTX_INPROC_SERVER = 0x1
_S_OK = 0


class _GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]


def _make_guid(guid_str: str) -> _GUID:
    d1, d2, d3, d4_01, d4_rest = guid_str.strip("{}").split("-")
    d4 = bytes([int(d4_01[i:i + 2], 16) for i in range(0, 4, 2)])
    d4 += bytes([int(d4_rest[i:i + 2], 16) for i in range(0, 12, 2)])
    return _GUID(
        int(d1, 16),
        int(d2, 16),
        int(d3, 16),
        (ctypes.c_ubyte * 8)(*d4),
    )


class TaskbarProgress:
    """任务栏进度条控制。用法：

        tb = TaskbarProgress(hwnd)
        tb.set_value(3, 10)          # 3/10，绿色
        tb.set_state(TBPF_ERROR)     # 红色
        tb.clear()                   # 结束清除
    """

    def __init__(self, hwnd: int = 0):
        self._hwnd = int(hwnd or 0)
        self._p = None  # ITaskbarList3 指针
        self._available = False
        if not sys.platform.startswith("win"):
            return
        try:
            self._init_com()
        except Exception as exc:
            logger.debug("任务栏进度条初始化失败(将静默降级): %s", exc)
            self._p = None
            self._available = False

    def _init_com(self) -> None:
        ole32 = ctypes.windll.ole32
        # CoInitialize 幂等；已初始化(返回 S_FALSE/RPC_E_CHANGED_MODE)忽略
        try:
            ole32.CoInitialize(None)
        except Exception:
            pass

        ptr = ctypes.c_void_p()
        clsid = _make_guid(_CLSID_TaskbarList)
        iid = _make_guid(_IID_ITaskbarList3)
        hr = ole32.CoCreateInstance(
            ctypes.byref(clsid),
            None,
            _CLSCTX_INPROC_SERVER,
            ctypes.byref(iid),
            ctypes.byref(ptr),
        )
        if hr != _S_OK or not ptr.value:
            raise OSError(f"CoCreateInstance 失败 hr=0x{hr & 0xFFFFFFFF:08X}")
        self._p = ptr

        vtable = ctypes.cast(
            ptr.value, ctypes.POINTER(ctypes.c_void_p)
        ).contents

        def vfunc(idx, restype, *argtypes):
            addr = vtable[idx]
            f = ctypes.WINFUNCTYPE(restype, ctypes.c_void_p, *argtypes)(addr)
            return f

        # ITaskbarList3 vtable 偏移
        self._HrInit = vfunc(3, ctypes.HRESULT)
        self._SetProgressValue = vfunc(
            9, ctypes.HRESULT, wintypes.HWND, ctypes.c_ulonglong, ctypes.c_ulonglong
        )
        self._SetProgressState = vfunc(10, ctypes.HRESULT, wintypes.HWND, ctypes.c_int)

        hr = self._HrInit(ptr)
        if hr != _S_OK:
            raise OSError(f"ITaskbarList3.HrInit 失败 hr=0x{hr & 0xFFFFFFFF:08X}")
        self._available = True

    @property
    def available(self) -> bool:
        return self._available

    def set_hwnd(self, hwnd: int) -> None:
        self._hwnd = int(hwnd or 0)

    def set_value(self, completed: int, total: int) -> bool:
        """更新进度值；自动进入绿色(NORMAL)状态。"""
        if not self._available or not self._hwnd or total <= 0:
            return False
        try:
            completed = max(0, min(int(completed), int(total)))
            self._SetProgressState(self._p, self._hwnd, TBPF_NORMAL)
            self._SetProgressValue(
                self._p, self._hwnd, ctypes.c_ulonglong(completed),
                ctypes.c_ulonglong(total),
            )
            return True
        except Exception as exc:
            logger.debug("任务栏 set_value 失败: %s", exc)
            return False

    def set_state(self, state: int) -> bool:
        """设置状态：NORMAL(绿)/ERROR(红)/PAUSED(黄)/INDETERMINATE(滚动)/NOPROGRESS(清除)。"""
        if not self._available or not self._hwnd:
            return False
        try:
            self._SetProgressState(self._p, self._hwnd, int(state))
            return True
        except Exception as exc:
            logger.debug("任务栏 set_state 失败: %s", exc)
            return False

    def set_indeterminate(self) -> bool:
        """不确定进度（初始化/准备阶段），任务栏显示滚动动画。"""
        return self.set_state(TBPF_INDETERMINATE)

    def set_error(self) -> bool:
        return self.set_state(TBPF_ERROR)

    def set_paused(self) -> bool:
        return self.set_state(TBPF_PAUSED)

    def clear(self) -> bool:
        """结束/空闲：清除任务栏进度条。"""
        return self.set_state(TBPF_NOPROGRESS)
