import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import signal
import sys
import tempfile
import traceback


APP_NAME = "ExcelSendWx"


def _get_log_path():
    candidates = []
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        candidates.append(Path(local_app_data) / APP_NAME / "logs")
    candidates.append(Path(tempfile.gettempdir()) / APP_NAME / "logs")

    for log_dir in candidates:
        try:
            log_dir.mkdir(parents=True, exist_ok=True)
            return log_dir / "app.log"
        except OSError:
            continue
    raise OSError("无法创建日志目录")


LOG_PATH = _get_log_path()
_LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"


def configure_logging():
    """配置 root logger；幂等，可在第三方库清空 handlers 后反复调用恢复。

    wechatauto 在 import 时会执行 ``root_logger.handlers.clear()``
    （wechatauto/logger.py 的 wxlog），会把本函数先前挂载的
    RotatingFileHandler 一并清掉。2026-08-29 起因该导入副作用，
    业务日志整整一个月只进控制台、app.log 仅剩启动行。因此 gui/wechatauto
    导入完成后必须再调用本函数一次，把文件 handler 挂回去。
    """
    root = logging.getLogger()
    target_path = os.path.normcase(os.path.abspath(str(LOG_PATH)))
    has_file_handler = any(
        isinstance(h, RotatingFileHandler)
        and os.path.normcase(os.path.abspath(getattr(h, "baseFilename", "")))
        == target_path
        for h in root.handlers
    )
    if not has_file_handler:
        file_handler = RotatingFileHandler(
            LOG_PATH,
            maxBytes=5 * 1024 * 1024,
            backupCount=3,
            encoding="utf-8",
        )
        file_handler.setFormatter(logging.Formatter(_LOG_FORMAT))
        root.addHandler(file_handler)

    # wechatauto 已自带一个控制台 StreamHandler，避免重复刷屏；只有在不存在
    # 非文件型 StreamHandler 时才补一个 stdout 处理器。
    # 注意 RotatingFileHandler 也是 StreamHandler 的子类，判断时要排除 FileHandler。
    if sys.stdout is not None and not any(
        isinstance(h, logging.StreamHandler)
        and not isinstance(h, logging.FileHandler)
        for h in root.handlers
    ):
        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setFormatter(logging.Formatter(_LOG_FORMAT))
        root.addHandler(stream_handler)

    # wechatauto 把 root 设为 DEBUG（同时把 comtypes 等提到 WARNING）。
    # 统一收到 INFO，避免第三方库 DEBUG 洪水把 app.log 撑爆轮转。
    root.setLevel(logging.INFO)


configure_logging()

logger = logging.getLogger(__name__)


def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.critical(
        "Unhandled exception",
        exc_info=(exc_type, exc_value, exc_traceback),
    )
    if sys.stdout is not None:
        print(f"程序发生未处理的异常，详细日志: {LOG_PATH}")


def handle_signal(signum, frame):
    logger.critical(
        "Received signal %s\n%s",
        signum,
        "".join(traceback.format_stack(frame)),
    )
    raise SystemExit(1)


sys.excepthook = handle_exception

try:
    signal.signal(signal.SIGINT, handle_signal)
    signal.signal(signal.SIGTERM, handle_signal)
except (AttributeError, OSError, ValueError):
    pass


def main():
    logger.info("Starting application; log file: %s", LOG_PATH)
    try:
        from modules.gui import run_gui
    except ImportError as exc:
        # gui 导入链中的 wechatauto 可能已清空 root handlers，先恢复再记录异常
        configure_logging()
        logger.exception("无法导入图形界面")
        if sys.stdout is not None:
            print(f"无法启动图形界面: {exc}")
            print("请安装依赖: pip install PyQt6 pandas openpyxl python-calamine wechatauto uiautomation requests")
        return 1

    # 关键：导入 modules.gui 会连带导入 wechatauto，其 import 时副作用
    # root_logger.handlers.clear() 会抹掉文件 handler，必须在此恢复，
    # 否则整个运行期（手动/定时/链路发送）业务日志都无法写入 app.log。
    configure_logging()
    logger.info("Logging reconfigured after GUI imports; log file: %s", LOG_PATH)

    run_gui()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
