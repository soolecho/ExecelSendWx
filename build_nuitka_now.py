# -*- coding: utf-8 -*-
"""
按 README 里的基线命令重新用 Nuitka 编译 main.py。
- 输出目录：dist_nuitka_config（与 installer.iss MyAppSourceDir 对齐）
- 图标：love.ico
- 关闭控制台
- 启用 PyQt6 / Tkinter 插件
- 加速：--jobs=进程数 --lto=no 首次更快（经验 439994）
- 全量输出写 logs\build_nuitka.log，便于排查
"""
import os
import sys
import subprocess
from datetime import datetime

PROJECT = os.path.dirname(os.path.abspath(__file__))
os.chdir(PROJECT)

LOG_DIR = os.path.join(PROJECT, "logs")
os.makedirs(LOG_DIR, exist_ok=True)
log_path = os.path.join(
    LOG_DIR, f"build_nuitka_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
)

env = os.environ.copy()
# 本地缓存目录，避免临时文件散落到 AppData，升级/切换时好清理。
# Nuitka 通过 env NUITKA_CACHE_DIR 读取缓存根目录（AppDirs._getCacheDir），
# ccache/bytecode 等子缓存都落在它下面；没有 CLI 参数 --cache-dir（4.1.3 不支持）。
N_CACHE = os.path.join(PROJECT, "nuitka_cache")
env["NUITKA_CACHE_DIR"] = N_CACHE
# 首次编译如果没装 C/C++ 编译器，让 Nuitka 静默自动下载 winlibs
env.setdefault("NUITKA_NON_INTERACTIVE", "1")

# 编译并发数：cc1 编译大模块时单进程峰值可达 2GB+，jobs 过高会在
# 多模块并行时 OOM（"cc1.exe: out of memory"）。按可用物理内存估算
# （每进程预留 2.5GB），并支持 NUITKA_JOBS 环境变量手动覆盖。
def _decide_jobs():
    forced = os.environ.get("NUITKA_JOBS", "").strip()
    cpu = os.cpu_count() or 4
    avail_gb = float(cpu)
    try:
        import ctypes

        class _MemStatus(ctypes.Structure):
            _fields_ = [
                ("dwLength", ctypes.c_ulong),
                ("dwMemoryLoad", ctypes.c_ulong),
                ("ullTotalPhys", ctypes.c_ulonglong),
                ("ullAvailPhys", ctypes.c_ulonglong),
                ("ullTotalPageFile", ctypes.c_ulonglong),
                ("ullAvailPageFile", ctypes.c_ulonglong),
                ("ullTotalVirtual", ctypes.c_ulonglong),
                ("ullAvailVirtual", ctypes.c_ulonglong),
                ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
            ]

        stat = _MemStatus()
        stat.dwLength = ctypes.sizeof(_MemStatus)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat))
        # 综合可用物理内存和可用提交内存，取较小值更保险
        avail_gb = min(stat.ullAvailPhys, stat.ullAvailPageFile) / (1024 ** 3)
    except Exception:
        pass
    if forced:
        try:
            return max(1, int(forced)), avail_gb
        except ValueError:
            pass
    by_mem = int(avail_gb // 2.5)
    return max(1, min(cpu, by_mem)), avail_gb

workers, free_gb = _decide_jobs()
print(f"[nuitka] jobs={workers} (cpu={os.cpu_count()}, "
      f"avail_mem~{free_gb:.1f}GB, override via NUITKA_JOBS)", flush=True)

cmd = [
    sys.executable, "-m", "nuitka",
    "--standalone",
    "--assume-yes-for-downloads",
    "--windows-icon-from-ico=love.ico",
    "--windows-console-mode=disable",
    "--enable-plugin=pyqt6",
    "--enable-plugin=tk-inter",
    f"--jobs={workers}",
    # LTO 关闭：开发期打包提速；如果要正式版更小更快可以再打开
    "--lto=no",
    # 编译产物缓存由 env NUITKA_CACHE_DIR 指定根目录（见上方 N_CACHE），
    # ccache/bytecode/dll-dependencies 子缓存自动落盘到该目录，后续增量编译
    # 只重编改动模块，显著加速迭代（4.1.3 无 --cache-dir CLI 选项）
    "--nofollow-import-to=wxauto4",
    # 函数内延迟导入的 Rust xlsx 引擎（大表格秒开，替代慢速 openpyxl），显式收进打包
    "--include-package=python_calamine",
    "--output-dir=dist_nuitka_config",
    "--output-filename=表格自动发送By春风予Lu.exe",
    "main.py",
]

print(f"[nuitka] cmd: {' '.join(cmd)}", flush=True)
print(f"[nuitka] log: {log_path}", flush=True)

with open(log_path, "w", encoding="utf-8") as log_fp:
    log_fp.write("COMMAND: " + " ".join(cmd) + "\n\n")
    log_fp.flush()
    p = subprocess.Popen(
        cmd,
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    assert p.stdout is not None
    for line in p.stdout:
        sys.stdout.write(line)
        sys.stdout.flush()
        log_fp.write(line)
        log_fp.flush()
    rc = p.wait()

print(f"\n[nuitka] exit code = {rc}", flush=True)
sys.exit(rc)
