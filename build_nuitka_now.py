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
from concurrent.futures import ProcessPoolExecutor

PROJECT = os.path.dirname(os.path.abspath(__file__))
os.chdir(PROJECT)

LOG_DIR = os.path.join(PROJECT, "logs")
os.makedirs(LOG_DIR, exist_ok=True)
log_path = os.path.join(
    LOG_DIR, f"build_nuitka_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
)

env = os.environ.copy()
# 本地缓存目录，避免临时文件散落到 AppData，升级/切换时好清理
env["NUITKA_CACHE_DIR"] = os.path.join(PROJECT, "nuitka_cache")
# 首次编译如果没装 C/C++ 编译器，让 Nuitka 静默自动下载 winlibs
env.setdefault("NUITKA_NON_INTERACTIVE", "1")

try:
    workers = ProcessPoolExecutor()._max_workers  # type: ignore[attr-defined]
except Exception:
    workers = max(2, (os.cpu_count() or 4) // 2)

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
