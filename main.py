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
log_handlers = [
    RotatingFileHandler(
        LOG_PATH,
        maxBytes=5 * 1024 * 1024,
        backupCount=3,
        encoding="utf-8",
    )
]
if sys.stdout is not None:
    log_handlers.append(logging.StreamHandler(sys.stdout))

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=log_handlers,
)

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
        logger.exception("无法导入图形界面")
        if sys.stdout is not None:
            print(f"无法启动图形界面: {exc}")
            print("请安装依赖: pip install PyQt6 pandas openpyxl python-calamine wxauto4 requests")
        return 1

    run_gui()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
