"""局域网网页控制服务（标准库实现，无第三方依赖）。

设计要点：
- HTTP 线程绝不直接操作 Qt 控件：所有涉及界面/任务的动作都通过
  ``WebBridge`` 投递到 Qt 主线程执行，并带超时等待结果；
- 上传与金山在线表格下载/解析在 HTTP 线程内完成（纯 Python，不碰 GUI），
  只有「启动发送/定时/监控」这类动作才回到主线程；
- 口令只保存 PBKDF2-SHA256 摘要与随机盐，明文不落盘；
- 会话使用 HttpOnly Cookie，默认 12 小时过期；
- 只监听局域网：绑定 0.0.0.0，但仅供内网访问（不要做公网端口映射）。
"""
import csv
import hashlib
import hmac
import json
import logging
import os
import re
import secrets
import socket
import tempfile
import threading
import time
from collections import deque
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any, Dict, Optional
from urllib.parse import parse_qs, urlparse

from PyQt6.QtCore import QObject, pyqtSignal

from . import web_ui

logger = logging.getLogger("ExcelSendWx.web")

PBKDF2_ITERATIONS = 200_000
SESSION_COOKIE = "wxweb_session"
SESSION_TTL_SEC = 12 * 3600
MAX_BODY_BYTES = 50 * 1024 * 1024  # 单次上传上限 50MB

DATA_EXTS = {".xlsx", ".xls", ".csv"}
ATTACH_EXTS = {
    ".xlsx", ".xls", ".csv", ".doc", ".docx", ".pdf", ".txt", ".zip",
    ".rar", ".7z", ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".pptx",
    ".ppt", ".mp4", ".mp3",
}

# 登录限速：同一 IP 60 秒内最多失败 8 次
LOGIN_FAIL_WINDOW = 60
LOGIN_FAIL_LIMIT = 8


def hash_password(password: str, salt_hex: str = "") -> tuple:
    """返回 (hash_hex, salt_hex)。salt_hex 为空时生成新盐。"""
    salt = bytes.fromhex(salt_hex) if salt_hex else secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", str(password).encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return digest.hex(), salt.hex()


def verify_password(password: str, hash_hex: str, salt_hex: str) -> bool:
    if not hash_hex or not salt_hex:
        return False
    try:
        digest, _ = hash_password(password, salt_hex)
    except ValueError:
        return False
    return hmac.compare_digest(digest, str(hash_hex))


def lan_ip() -> str:
    """取本机在局域网中的 IPv4 地址（无网络时回退 127.0.0.1）。"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        sock.connect(("8.8.8.8", 80))
        return sock.getsockname()[0]
    except OSError:
        try:
            return socket.gethostbyname(socket.gethostname())
        except OSError:
            return "127.0.0.1"
    finally:
        sock.close()


class WebError(Exception):
    """带 HTTP 状态码的业务异常。"""

    def __init__(self, message: str, status: int = 400):
        super().__init__(message)
        self.status = status


class WebStopping(WebError):
    def __init__(self):
        super().__init__("服务正在停止，请稍后重试", 503)


class WebTimeout(WebError):
    def __init__(self):
        super().__init__("主程序响应超时，请稍后重试", 504)


class WebBridge(QObject):
    """把可调用对象安全地投递到 Qt 主线程执行，并阻塞等待结果。"""

    _dispatch = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._lock = threading.Lock()
        self._pending: Dict[str, dict] = {}
        self.stopping = False
        self._dispatch.connect(self._run_on_main)

    def call(self, func, timeout: float = 15.0):
        if self.stopping:
            raise WebStopping()
        req_id = secrets.token_hex(8)
        event = threading.Event()
        box = {"event": event, "result": None, "error": None, "func": func}
        with self._lock:
            self._pending[req_id] = box
        self._dispatch.emit(req_id)
        if not event.wait(max(1.0, float(timeout))):
            with self._lock:
                self._pending.pop(req_id, None)
            raise WebTimeout()
        if box["error"] is not None:
            raise box["error"]
        return box["result"]

    def _run_on_main(self, req_id: str):
        with self._lock:
            box = self._pending.pop(req_id, None)
        if box is None:
            return
        try:
            box["result"] = box["func"]()
        except Exception as exc:  # 异常原样带回 HTTP 线程
            box["error"] = exc
        finally:
            box["event"].set()

    def shutdown(self):
        self.stopping = True
        with self._lock:
            boxes = list(self._pending.values())
            self._pending.clear()
        for box in boxes:
            box["error"] = WebStopping()
            box["event"].set()


class WebActions:
    """主线程侧的网页动作实现（由 MainWindow 构造，方法只在主线程调用）。"""

    def __init__(self, window):
        self.window = window
        self.jobs: Dict[str, dict] = {}
        self._job_seq = 0

    # ---------------- 状态 / 日志 ----------------
    def status(self) -> dict:
        w = self.window
        manual_busy = False
        try:
            worker = w.table_filter_tab.worker
            manual_busy = bool(worker and worker.isRunning())
        except Exception:
            pass
        schedule_busy = False
        try:
            sw = w.schedule_tab.send_worker
            schedule_busy = bool(sw and sw.isRunning())
        except Exception:
            pass
        monitor_busy = False
        try:
            mw = w.monitor_tab._one_shot_worker
            monitor_busy = bool(mw and mw.isRunning())
        except Exception:
            pass
        return {
            "app": getattr(w, "windowTitle", lambda: "")() or "",
            "busy": manual_busy or schedule_busy,
            "manual_busy": manual_busy,
            "schedule_busy": schedule_busy,
            "monitor_busy": monitor_busy,
        }

    def progress(self) -> dict:
        """当前发送进度（复用底部坞进度条，手动/定时/配置链都可见）。"""
        w = self.window
        active = bool(getattr(w, "_send_progress_active", False))
        current = 0
        total = 0
        text = ""
        dock = getattr(w, "bottom_dock", None)
        if dock is not None:
            try:
                bar = dock.progress_bar
                maximum = int(bar.maximum())
                if maximum > 0:
                    total = maximum
                    current = max(0, min(int(bar.value()), total))
                text = str(dock.progress_label.text())
            except Exception:
                pass
        percent = int(current * 100 / total) if total > 0 else 0
        return {
            "active": active,
            "current": current,
            "total": total,
            "percent": percent,
            "text": text,
            "owner": str(getattr(w, "_progress_owner", "") or ""),
        }

    def logs(self, since: int = 0, limit: int = 200) -> dict:
        w = self.window
        buffer = getattr(w, "_global_log_buffer", None)
        if buffer is None:
            return {"items": [], "last": int(since)}
        try:
            since = int(since)
        except (TypeError, ValueError):
            since = 0
        try:
            limit = max(1, min(int(limit), 500))
        except (TypeError, ValueError):
            limit = 200
        with getattr(w, "_global_log_lock", threading.Lock()):
            items = [dict(x) for x in buffer if int(x.get("seq", 0)) > since]
        if len(items) > limit:
            items = items[-limit:]
        last = items[-1]["seq"] if items else since
        return {"items": items, "last": last}

    # ---------------- 电脑端已保存配置（数据发送） ----------------
    def list_profiles(self) -> dict:
        items = []
        for item in self._profile_entries():
            safe = dict(item)
            safe.pop("path", None)
            items.append(safe)
        return {"items": items}

    def _profile_entries(self):
        """汇总「最近配置」+ 配置目录下所有 *.json（只读，不修改任何文件）。"""
        cm = self.window.config_manager
        paths = []
        try:
            paths.extend(cm.get_recent_profiles())
        except Exception:
            pass
        try:
            for p in sorted(cm.profile_dir.glob("*.json")):
                paths.append(str(p))
        except Exception:
            pass
        items = []
        seen = set()
        for raw_path in paths:
            path = os.path.abspath(str(raw_path))
            if path in seen or not os.path.isfile(path):
                continue
            seen.add(path)
            try:
                with open(path, "r", encoding="utf-8") as fp:
                    raw = json.load(fp)
            except Exception:
                continue
            if not isinstance(raw, dict):
                continue
            excel = raw.get("excel") or {}
            send = raw.get("send") or {}
            if str(excel.get("source", "local")) == "wps_cloud":
                file_label = "云文档: " + str(
                    excel.get("cloud_file_name")
                    or excel.get("cloud_file_id") or "")
            else:
                file_label = os.path.basename(str(excel.get("path", "") or ""))
            columns = excel.get("extract_columns") or []
            items.append({
                "id": hashlib.sha1(path.encode("utf-8")).hexdigest()[:16],
                # path 仅供服务端定位文件，返回给浏览器前会剔除
                "path": path,
                "name": os.path.basename(path),
                "saved_at": str(raw.get("saved_at", "") or ""),
                "excel": file_label,
                "sheet": str(excel.get("sheet", "") or ""),
                "name_column": str(excel.get("name_column", "") or ""),
                "columns": len(columns) if isinstance(columns, list) else 0,
                "recipients": str(send.get("manual_recipient", "") or ""),
                "auto_send": bool(send.get("auto_send", False)),
                "interval": send.get("interval", 0.5),
                "agent_api": bool(send.get("agent_api", False)),
            })
        items.sort(key=lambda x: x.get("saved_at", ""), reverse=True)
        return items

    def _resolve_profile(self, profile_id: str):
        target = str(profile_id or "").strip()
        if not target:
            return None
        for item in self._profile_entries():
            if item["id"] == target:
                return item
        return None

    def run_profile(self, profile_id: str) -> dict:
        item = self._resolve_profile(profile_id)
        if item is None:
            return {"ok": False, "error": "配置不存在或已被删除"}
        tab = self.window.table_filter_tab
        result = tab.web_run_profile(item["path"], trigger_auto_send=True)
        if result.get("ok"):
            result["name"] = item["name"]
        return result

    def edit_profile(self, profile_id: str) -> dict:
        item = self._resolve_profile(profile_id)
        if item is None:
            return {"ok": False, "error": "配置不存在或已被删除"}
        tab = self.window.table_filter_tab
        result = tab.web_run_profile(item["path"], trigger_auto_send=False)
        if result.get("ok"):
            result["name"] = item["name"]
            try:
                self.window.show_main_window()
                self.window._on_nav_page("send")
            except Exception:
                pass
        return result

    def get_profile_detail(self, profile_id: str) -> dict:
        """读取配置内容，供网页表单编辑。"""
        item = self._resolve_profile(profile_id)
        if item is None:
            return {"ok": False, "error": "配置不存在或已被删除"}
        try:
            with open(item["path"], "r", encoding="utf-8") as fp:
                raw = json.load(fp)
        except Exception as exc:
            return {"ok": False, "error": f"读取配置失败：{exc}"}
        excel = raw.get("excel") or {}
        send = raw.get("send") or {}
        columns = excel.get("extract_columns") or []
        if not isinstance(columns, list):
            columns = []
        return {
            "ok": True,
            "profile": {
                "id": item["id"],
                "name": item["name"],
                "saved_at": str(raw.get("saved_at", "") or ""),
                "source": str(excel.get("source", "local") or "local"),
                "excel_path": str(excel.get("path", "") or ""),
                "cloud_file_id": str(excel.get("cloud_file_id", "") or ""),
                "cloud_file_name": str(excel.get("cloud_file_name", "") or ""),
                "sheet": str(excel.get("sheet", "") or ""),
                "header_row": int(excel.get("header_row", 1) or 1),
                "name_column": str(excel.get("name_column", "") or ""),
                "extract_columns": [str(c) for c in columns],
                "wechat_column": str(excel.get("wechat_column", "") or ""),
                "manual_recipient": str(
                    send.get("manual_recipient", "") or ""),
                "custom_message_enabled": bool(
                    send.get("custom_message_enabled", False)),
                "custom_message": str(send.get("custom_message", "") or ""),
                "auto_send": bool(send.get("auto_send", False)),
                "interval": float(send.get("interval", 0.5) or 0.5),
                "chat_delay": float(send.get("chat_delay", 0.3) or 0.3),
                "send_order": list(send.get("send_order") or []),
                "attachment": str(send.get("attachment", "") or ""),
                "region_range": str(send.get("region_range", "") or ""),
                "region_sheet": str(send.get("region_sheet", "") or ""),
                "at_enabled": bool(send.get("at_enabled", False)),
                "at_members": str(send.get("at_members", "") or ""),
                "at_all": bool(send.get("at_all", False)),
                "at_attach": str(send.get("at_attach", "text") or "text"),
                "agent_api": bool(send.get("agent_api", False)),
            },
        }

    def save_profile(self, profile_id: str, data: dict) -> dict:
        """把网页表单的修改写回原配置 JSON（保留筛选条件/联系人映射等未编辑字段）。"""
        from .config_manager import ConfigError
        item = self._resolve_profile(profile_id)
        if item is None:
            return {"ok": False, "error": "配置不存在或已被删除"}
        try:
            with open(item["path"], "r", encoding="utf-8") as fp:
                raw = json.load(fp)
        except Exception as exc:
            return {"ok": False, "error": f"读取配置失败：{exc}"}
        if not isinstance(raw, dict):
            return {"ok": False, "error": "配置内容格式不正确"}

        def _text(key, default=""):
            value = data.get(key, default)
            if value is None:
                return default
            return str(value).strip()

        excel = raw.setdefault("excel", {})
        send = raw.setdefault("send", {})

        source = _text("source", "local") or "local"
        if source not in ("local", "wps_cloud"):
            source = "local"
        excel["source"] = source
        excel["sheet"] = _text("sheet")
        excel["name_column"] = _text("name_column")
        excel["wechat_column"] = _text("wechat_column")
        try:
            excel["header_row"] = max(1, int(data.get("header_row", 1) or 1))
        except (TypeError, ValueError):
            excel["header_row"] = 1

        raw_columns = data.get("extract_columns")
        if isinstance(raw_columns, str):
            raw_columns = re.split(r"[,，;；\n]+", raw_columns)
        columns = []
        if isinstance(raw_columns, list):
            for col in raw_columns:
                text = str(col).strip()
                if text and text not in columns:
                    columns.append(text)
        excel["extract_columns"] = columns

        if source == "wps_cloud":
            link = _text("cloud_file_id")
            if link:
                try:
                    from .wps_cloud import WpsCloudError, parse_file_input
                    excel["cloud_file_id"] = parse_file_input(link)
                except WpsCloudError as exc:
                    return {"ok": False, "error": f"云文档链接无效：{exc}"}
            excel["cloud_file_name"] = _text("cloud_file_name")
            excel["path"] = ""
        else:
            excel["path"] = _text("excel_path")
            if not excel["path"]:
                return {"ok": False, "error": "请填写 Excel 文件路径"}

        send["manual_recipient"] = _text("manual_recipient")
        send["custom_message_enabled"] = bool(
            data.get("custom_message_enabled", False))
        send["custom_message"] = str(data.get("custom_message", "") or "")
        send["auto_send"] = bool(data.get("auto_send", False))
        try:
            send["interval"] = float(data.get("interval", 0.5) or 0.0)
            send["chat_delay"] = float(data.get("chat_delay", 0.3) or 0.0)
        except (TypeError, ValueError):
            send["interval"] = 0.5
            send["chat_delay"] = 0.3
        send["attachment"] = _text("attachment")
        send["region_range"] = _text("region_range")
        send["region_sheet"] = _text("region_sheet")
        send["at_enabled"] = bool(data.get("at_enabled", False))
        send["at_members"] = _text("at_members")
        send["at_all"] = bool(data.get("at_all", False))
        attach = _text("at_attach", "text") or "text"
        send["at_attach"] = "first" if attach == "first" else "text"
        send["agent_api"] = bool(data.get("agent_api", False))

        order = data.get("send_order")
        valid_order = {"text", "image", "region", "custom", "attachment"}
        if isinstance(order, list):
            cleaned = []
            for key in order:
                key = str(key).strip()
                if key in valid_order and key not in cleaned:
                    cleaned.append(key)
            if cleaned:
                send["send_order"] = cleaned

        # 发送形式兼容字段：至少保证 text/image 组合有值
        if "send_order" not in send:
            send["mode"] = "text"

        try:
            self.window.config_manager.save_profile(item["path"], raw)
        except ConfigError as exc:
            return {"ok": False, "error": str(exc)}
        except Exception as exc:
            return {"ok": False, "error": f"保存失败：{exc}"}
        try:
            self.window.table_filter_tab.refresh_recent_configs()
        except Exception:
            pass
        return {"ok": True, "message": "配置已保存"}

    # ---------------- 手动发送 ----------------
    def send(self, payload: dict, grid: list, file_path: str) -> dict:
        tab = self.window.table_filter_tab
        result = tab.web_send(payload, grid, file_path)
        if not result.get("ok"):
            return result
        worker = result.get("worker")
        job_id = self._new_job()
        job = {
            "id": job_id,
            "kind": "send",
            "status": "running",
            "total": int(result.get("total", 0) or 0),
            "current": 0,
            "success": 0,
            "failed": 0,
            "message": "正在发送",
            "started": time.time(),
        }
        self.jobs[job_id] = job

        def _progress(current, total):
            job["current"] = int(current)
            job["total"] = int(total)
            job["message"] = f"发送进度 {current}/{total}"

        def _result(res):
            try:
                success, failed, total, _failed_tasks = res
                job["success"] = int(success)
                job["failed"] = int(failed)
                job["total"] = int(total)
                job["status"] = "done"
                job["message"] = f"完成：成功 {success}，失败 {failed}"
            except Exception:
                job["status"] = "done"
                job["message"] = "发送完成"

        def _error(err):
            job["status"] = "error"
            job["message"] = f"发送出错：{err}"

        def _finished():
            if job["status"] == "running":
                job["status"] = "done"
                job["message"] = "发送已结束"

        try:
            worker.signals.progress.connect(_progress)
            worker.signals.result.connect(_result)
            worker.signals.error.connect(_error)
            worker.finished.connect(_finished)
        except Exception:
            pass
        # 极端情况下（任务极小）连接信号前 worker 已结束，避免任务永远停在 running
        try:
            if not worker.isRunning() and job["status"] == "running":
                job["status"] = "done"
                job["message"] = "发送已结束"
        except Exception:
            pass
        self._trim_jobs()
        return {"ok": True, "job_id": job_id, "total": job["total"]}

    def job(self, job_id: str) -> dict:
        job = self.jobs.get(str(job_id))
        if job is None:
            return {"ok": False, "error": "任务不存在"}
        return {"ok": True, "job": dict(job)}

    def _new_job(self) -> str:
        self._job_seq += 1
        return f"job-{int(time.time())}-{self._job_seq}"

    def _trim_jobs(self):
        if len(self.jobs) <= 50:
            return
        for key in list(self.jobs.keys())[:-50]:
            self.jobs.pop(key, None)

    # ---------------- 定时任务 ----------------
    def list_schedules(self) -> dict:
        store = self.window.schedule_store
        items = []
        for task in store.load_all():
            rule = task.repeat_mode
            if rule == "daily":
                rule_text = "每天"
            elif rule == "weekly":
                from .schedule_manager import WEEKDAY_NAMES
                rule_text = "每周 " + "、".join(
                    WEEKDAY_NAMES[d - 1] for d in task.days if 1 <= d <= 7)
            else:
                rule_text = "指定日期 " + "、".join(task.run_dates or [])
            if getattr(task, "kind", "message") == "profiles":
                target = f"配置链 {len(task.profile_paths or [])} 个"
            else:
                target = f"接收人 {len(task.recipients or [])} 个"
            items.append({
                "id": task.id,
                "name": task.name,
                "enabled": bool(task.enabled),
                "kind": getattr(task, "kind", "message"),
                "rule": rule_text,
                "times": list(task.times or []),
                "target": target,
                "message": (task.message or "")[:60],
            })
        return {"items": items}

    def run_schedule(self, task_id: str) -> dict:
        tasks = {t.id: t for t in self.window.schedule_store.load_all()}
        task = tasks.get(str(task_id))
        if task is None:
            return {"ok": False, "error": "任务不存在"}
        ok, message = self.window.schedule_tab.start_task_now(task)
        return {"ok": bool(ok), "message": message,
                "error": "" if ok else message}

    def set_schedule_enabled(self, task_id: str, enabled: bool) -> dict:
        store = self.window.schedule_store
        tasks = store.load_all()
        target = None
        for task in tasks:
            if task.id == str(task_id):
                target = task
                break
        if target is None:
            return {"ok": False, "error": "任务不存在"}
        target.enabled = bool(enabled)
        store.save(target)
        try:
            self.window.schedule_tab.reload_tasks()
        except Exception:
            pass
        return {"ok": True, "enabled": bool(enabled)}

    # ---------------- 监控任务 ----------------
    def list_monitors(self) -> dict:
        items = []
        for task in self.window.monitor_store.load_all():
            rule = {"weekly": "每周", "once": "一次"}.get(
                task.repeat_mode, "每天")
            items.append({
                "id": task.id,
                "name": task.name,
                "enabled": bool(task.enabled),
                "rule": rule,
                "path": getattr(task, "watch_path", "") or "",
                "interval": int(getattr(task, "poll_interval_min", 0) or 0),
                "target": f"接收人 {len(task.recipients or [])} 个",
            })
        return {"items": items}

    def run_monitor(self, task_id: str) -> dict:
        ok, message = self.window.monitor_tab.start_task_once(str(task_id))
        return {"ok": bool(ok), "message": message,
                "error": "" if ok else message}

    def set_monitor_enabled(self, task_id: str, enabled: bool) -> dict:
        ok = self.window.monitor_store.set_enabled(str(task_id), bool(enabled))
        if not ok:
            return {"ok": False, "error": "任务不存在"}
        try:
            self.window.monitor_tab.reload_tasks()
        except Exception:
            pass
        return {"ok": True, "enabled": bool(enabled)}


class WebService:
    """内置 HTTP 服务：负责端口监听、登录鉴权、上传与路由。"""

    def __init__(self, bridge: WebBridge, config_manager, actions: WebActions):
        self.bridge = bridge
        self.config_manager = config_manager
        self.actions = actions
        self._server: Optional[ThreadingHTTPServer] = None
        self._thread: Optional[threading.Thread] = None
        self._port = 0
        self._lock = threading.Lock()
        # 会话持久化：重启软件后浏览器 Cookie 仍有效（12 小时内免登录）
        self._sess_lock = threading.Lock()
        self._sessions_file = None
        base_dir = getattr(config_manager, "base_dir", None)
        if base_dir:
            self._sessions_file = os.path.join(
                str(base_dir), "web_sessions.json")
        self._sessions: Dict[str, float] = self._load_sessions()
        self._uploads: Dict[str, dict] = {}
        self._upload_lock = threading.Lock()
        self._login_fails: Dict[str, deque] = {}
        self._login_lock = threading.Lock()
        self._upload_dir = os.path.join(
            os.environ.get("LOCALAPPDATA")
            or tempfile.gettempdir(), "ExcelSendWx", "web_uploads")
        try:
            os.makedirs(self._upload_dir, exist_ok=True)
        except OSError:
            self._upload_dir = tempfile.mkdtemp(prefix="esw_web_")

    # ---------------- 生命周期 ----------------
    @property
    def running(self) -> bool:
        return self._server is not None

    @property
    def port(self) -> int:
        return self._port

    def start(self, port: int) -> None:
        with self._lock:
            if self._server is not None:
                return
            server = _WebHTTPServer(("0.0.0.0", int(port)), _WebRequestHandler)
            server.service = self
            server.daemon_threads = True
            thread = threading.Thread(
                target=server.serve_forever, daemon=True,
                name="ExcelSendWxWeb")
            thread.start()
            self._server = server
            self._thread = thread
            self._port = int(port)
            logger.info("网页控制服务已启动: http://%s:%s", lan_ip(), port)

    def stop(self) -> None:
        with self._lock:
            server, thread = self._server, self._thread
            self._server = None
            self._thread = None
            self._port = 0
        if server is not None:
            try:
                server.shutdown()
            except Exception:
                pass
            try:
                server.server_close()
            except Exception:
                pass
        if thread is not None and thread.is_alive():
            thread.join(timeout=2.0)

    def restart(self, port: int) -> None:
        self.stop()
        self.start(port)

    def url(self) -> str:
        return f"http://{lan_ip()}:{self._port}/"

    # ---------------- 会话 ----------------
    def _sessions_lock(self):
        lock = getattr(self, "_sess_lock", None)
        if lock is None:
            lock = threading.Lock()
            self._sess_lock = lock
        return lock

    def _load_sessions(self) -> Dict[str, float]:
        if not self._sessions_file:
            return {}
        try:
            with open(self._sessions_file, "r", encoding="utf-8") as fp:
                data = json.load(fp)
        except (OSError, ValueError):
            return {}
        if not isinstance(data, dict):
            return {}
        now = time.time()
        result = {}
        for token, expiry in data.items():
            try:
                expiry = float(expiry)
            except (TypeError, ValueError):
                continue
            if expiry > now:
                result[str(token)] = expiry
        if result:
            logger.info("[web] 已恢复 %d 个登录会话（免登录）", len(result))
        return result

    def _save_sessions(self):
        sessions_file = getattr(self, "_sessions_file", None)
        if not sessions_file:
            return
        data = dict(self._sessions)
        try:
            tmp = sessions_file + ".tmp"
            with open(tmp, "w", encoding="utf-8") as fp:
                json.dump(data, fp, ensure_ascii=False)
            os.replace(tmp, sessions_file)
        except OSError as exc:
            logger.warning("[web] 保存登录会话失败: %s", exc)

    def create_session(self) -> str:
        token = secrets.token_urlsafe(32)
        with self._sessions_lock():
            self._sessions[token] = time.time() + SESSION_TTL_SEC
            self._save_sessions()
        return token

    def check_session(self, token: str) -> bool:
        if not token:
            return False
        now = time.time()
        with self._sessions_lock():
            expiry = self._sessions.get(token)
            if expiry is None:
                return False
            if expiry < now:
                self._sessions.pop(token, None)
                self._save_sessions()
                return False
            return True

    def drop_session(self, token: str) -> None:
        with self._sessions_lock():
            if self._sessions.pop(token, None) is not None:
                self._save_sessions()

    def revoke_all_sessions(self) -> None:
        """修改密码后让所有设备重新登录。"""
        with self._sessions_lock():
            self._sessions.clear()
            self._save_sessions()

    def login_allowed(self, ip: str) -> bool:
        now = time.time()
        with self._login_lock:
            bucket = self._login_fails.setdefault(ip, deque())
            while bucket and now - bucket[0] > LOGIN_FAIL_WINDOW:
                bucket.popleft()
            return len(bucket) < LOGIN_FAIL_LIMIT

    def note_login_fail(self, ip: str) -> None:
        with self._login_lock:
            self._login_fails.setdefault(ip, deque()).append(time.time())

    def clear_login_fails(self, ip: str) -> None:
        with self._login_lock:
            self._login_fails.pop(ip, None)

    # ---------------- 上传 ----------------
    def save_upload(self, filename: str, body: bytes, kind: str = "data") -> str:
        name = os.path.basename(str(filename or "")).strip() or "upload.bin"
        ext = os.path.splitext(name)[1].lower()
        allowed = DATA_EXTS if kind != "attachment" else ATTACH_EXTS
        if ext not in allowed:
            raise WebError(f"不支持的文件类型：{ext or '未知'}")
        upload_id = secrets.token_hex(12)
        safe_name = re.sub(r"[^\w.\-\u4e00-\u9fff]+", "_", name)
        dest = os.path.join(self._upload_dir, f"{upload_id}_{safe_name}")
        try:
            with open(dest, "wb") as fp:
                fp.write(body)
        except OSError as exc:
            raise WebError(f"保存文件失败：{exc}", 500)
        with self._upload_lock:
            self._uploads[upload_id] = {
                "path": dest,
                "name": name,
                "kind": kind,
                "created": time.time(),
            }
            self._cleanup_uploads_locked()
        return upload_id

    def get_upload(self, upload_id: str) -> Optional[dict]:
        with self._upload_lock:
            info = self._uploads.get(str(upload_id))
            return dict(info) if info else None

    def _cleanup_uploads_locked(self):
        limit = time.time() - 6 * 3600
        for key in list(self._uploads.keys()):
            info = self._uploads[key]
            if info.get("created", 0) < limit:
                self._uploads.pop(key, None)
                try:
                    os.remove(info.get("path", ""))
                except OSError:
                    pass

    def cleanup_all_uploads(self) -> None:
        with self._upload_lock:
            uploads = list(self._uploads.values())
            self._uploads.clear()
        for info in uploads:
            try:
                os.remove(info.get("path", ""))
            except OSError:
                pass

    # ---------------- 表格解析 ----------------
    @staticmethod
    def read_sheet_names(path: str) -> list:
        if path.lower().endswith(".csv"):
            return ["CSV"]
        from .wps_cloud import read_sheet_names
        return list(read_sheet_names(path) or [])

    @staticmethod
    def read_grid(path: str, sheet: str = "") -> list:
        if path.lower().endswith(".csv"):
            rows = []
            with open(path, "r", encoding="utf-8-sig", errors="replace",
                      newline="") as fp:
                for row in csv.reader(fp):
                    rows.append(["" if c is None else c for c in row])
            return rows
        from .wps_cloud import read_one_sheet
        names = WebService.read_sheet_names(path)
        target = sheet or (names[0] if names else "")
        if target and target not in names:
            raise WebError(f"工作表不存在：{target}")
        grid = read_one_sheet(path, target)
        if grid is None:
            raise WebError(f"工作表读取失败：{target}")
        return grid

    @staticmethod
    def build_preview(grid: list, header_row: int = 1) -> dict:
        from .table_processor import TableProcessor
        if not grid:
            raise WebError("表格内容为空")
        try:
            header_row = max(1, min(int(header_row), len(grid)))
        except (TypeError, ValueError):
            header_row = 1
        headers = TableProcessor.derive_headers(grid[header_row - 1])
        rows = grid[header_row:]
        preview = []
        for row in rows[:5]:
            preview.append(["" if c is None else str(c) for c in row])
        return {
            "headers": [str(h) for h in headers],
            "preview": preview,
            "total_rows": max(0, len(rows)),
            "header_row": header_row,
        }

    # ---------------- 登录 ----------------
    def verify_login(self, password: str, ip: str) -> str:
        settings = self.config_manager.load_global_settings()
        hash_hex = str(settings.get("web_password_hash", "") or "")
        salt_hex = str(settings.get("web_password_salt", "") or "")
        if not hash_hex or not salt_hex:
            raise WebError("管理员尚未设置网页登录密码，请先在电脑端「全量设置」里设置")
        if not self.login_allowed(ip):
            logger.warning("[web] 登录被限速 ip=%s", ip)
            raise WebError("尝试次数过多，请 1 分钟后再试", 429)
        # 保存与登录都去掉首尾空白，避免手机键盘 / 复制粘贴带入空格导致对不上
        password = str(password or "").strip()
        if not verify_password(password, hash_hex, salt_hex):
            self.note_login_fail(ip)
            logger.warning("[web] 登录失败 ip=%s 原因=密码不匹配", ip)
            raise WebError("密码错误", 401)
        self.clear_login_fails(ip)
        logger.info("[web] 登录成功 ip=%s", ip)
        return self.create_session()

    def service_info(self) -> dict:
        return {"ok": True, "service": "ExcelSendWx", "port": self._port}


class _WebHTTPServer(ThreadingHTTPServer):
    # Windows 下 SO_REUSEADDR 会允许「抢占」已被占用的端口，
    # 这里关闭它，端口冲突时明确报错而不是两个服务同时监听。
    allow_reuse_address = False
    daemon_threads = True

    def __init__(self, addr, handler):
        try:
            super().__init__(addr, handler)
        except OSError as exc:
            raise WebError(f"端口被占用或无法监听：{exc}", 500)


class _WebRequestHandler(BaseHTTPRequestHandler):
    server_version = "ExcelSendWxWeb/1.0"
    protocol_version = "HTTP/1.1"

    # ---------------- 基础输出 ----------------
    def log_message(self, fmt, *args):
        logger.debug("[web] " + fmt, *args)

    @property
    def service(self) -> WebService:
        return self.server.service

    def _send_json(self, data: dict, status: int = 200, cookie: str = ""):
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        # 错误响应未消费请求体时关闭连接，避免 keep-alive 下一条请求解析错位
        if status >= 400:
            self.close_connection = True
            self.send_header("Connection", "close")
        if cookie:
            self.send_header("Set-Cookie", cookie)
        self.end_headers()
        try:
            self.wfile.write(payload)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _send_html(self, text: str):
        payload = text.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try:
            self.wfile.write(payload)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _read_body(self, limit: int = 1024 * 1024) -> bytes:
        try:
            length = int(self.headers.get("Content-Length", "0") or 0)
        except ValueError:
            length = 0
        if length <= 0:
            return b""
        if length > limit:
            raise WebError("请求内容过大", 413)
        return self.rfile.read(length)

    def _read_json(self) -> dict:
        body = self._read_body()
        if not body:
            return {}
        try:
            data = json.loads(body.decode("utf-8"))
        except (ValueError, UnicodeDecodeError) as exc:
            raise WebError(f"请求格式错误：{exc}")
        return data if isinstance(data, dict) else {}

    def _cookie_token(self) -> str:
        raw = self.headers.get("Cookie", "") or ""
        for part in raw.split(";"):
            key, _, value = part.strip().partition("=")
            if key == SESSION_COOKIE:
                return value.strip()
        return ""

    def _require_auth(self) -> bool:
        if self.service.check_session(self._cookie_token()):
            return True
        self._send_json(
            {"ok": False, "error": "未登录或登录已过期", "need_login": True},
            status=401)
        return False

    def _client_ip(self) -> str:
        try:
            return self.client_address[0]
        except Exception:
            return "?"

    # ---------------- GET ----------------
    def do_GET(self):
        try:
            path = urlparse(self.path).path
            if path in ("/", "/index.html"):
                self._send_html(web_ui.PAGE_HTML)
                return
            if path == "/api/service":
                self._send_json(self.service.service_info())
                return
            if not path.startswith("/api/"):
                self._send_json({"ok": False, "error": "Not Found"}, 404)
                return
            if not self._require_auth():
                return
            query = parse_qs(urlparse(self.path).query)
            if path == "/api/status":
                data = self.service.bridge.call(
                    self.service.actions.status, timeout=10)
                data.update({
                    "ok": True,
                    "lan_url": self.service.url(),
                    "port": self.service.port,
                })
                self._send_json(data)
                return
            if path == "/api/logs":
                since = int((query.get("since") or ["0"])[0] or 0)
                limit = int((query.get("limit") or ["200"])[0] or 200)
                data = self.service.bridge.call(
                    lambda: self.service.actions.logs(since, limit), timeout=10)
                data["ok"] = True
                self._send_json(data)
                return
            if path == "/api/progress":
                data = self.service.bridge.call(
                    self.service.actions.progress, timeout=10)
                data["ok"] = True
                self._send_json(data)
                return
            if path == "/api/schedule":
                data = self.service.bridge.call(
                    self.service.actions.list_schedules, timeout=10)
                data["ok"] = True
                self._send_json(data)
                return
            if path == "/api/profiles":
                data = self.service.bridge.call(
                    self.service.actions.list_profiles, timeout=10)
                data["ok"] = True
                self._send_json(data)
                return
            if path == "/api/profiles/detail":
                pid = (query.get("id") or [""])[0]
                data = self.service.bridge.call(
                    lambda: self.service.actions.get_profile_detail(pid),
                    timeout=10)
                self._send_json(data)
                return
            if path == "/api/monitor":
                data = self.service.bridge.call(
                    self.service.actions.list_monitors, timeout=10)
                data["ok"] = True
                self._send_json(data)
                return
            if path.startswith("/api/jobs/"):
                job_id = path[len("/api/jobs/"):]
                data = self.service.bridge.call(
                    lambda: self.service.actions.job(job_id), timeout=10)
                self._send_json(data)
                return
            self._send_json({"ok": False, "error": "未知接口"}, 404)
        except WebError as exc:
            self._send_json({"ok": False, "error": str(exc)}, exc.status)
        except Exception as exc:
            logger.exception("web GET 处理异常")
            self._send_json({"ok": False, "error": f"服务器错误：{exc}"}, 500)

    # ---------------- POST ----------------
    def do_POST(self):
        try:
            path = urlparse(self.path).path
            query = parse_qs(urlparse(self.path).query)
            if path == "/api/login":
                self._handle_login()
                return
            if not path.startswith("/api/"):
                self._send_json({"ok": False, "error": "Not Found"}, 404)
                return
            if not self._require_auth():
                return
            if path == "/api/logout":
                self.service.drop_session(self._cookie_token())
                self._send_json(
                    {"ok": True},
                    cookie=f"{SESSION_COOKIE}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax")
                return
            if path == "/api/upload":
                self._handle_upload(query)
                return
            if path == "/api/preview":
                self._handle_preview()
                return
            if path == "/api/names":
                self._handle_names()
                return
            if path == "/api/cloud/preview":
                self._handle_cloud_preview()
                return
            if path == "/api/send":
                self._handle_send()
                return
            if path == "/api/profiles/run":
                body = self._read_json()
                data = self.service.bridge.call(
                    lambda: self.service.actions.run_profile(
                        str(body.get("id", "") or "")), timeout=20)
                self._send_json(data)
                return
            if path == "/api/profiles/edit":
                body = self._read_json()
                data = self.service.bridge.call(
                    lambda: self.service.actions.edit_profile(
                        str(body.get("id", "") or "")), timeout=20)
                self._send_json(data)
                return
            if path == "/api/profiles/save":
                body = self._read_json()
                pid = str(body.get("id", "") or "")
                data = self.service.bridge.call(
                    lambda: self.service.actions.save_profile(pid, body),
                    timeout=15)
                self._send_json(data)
                return
            if path.startswith("/api/schedule/"):
                self._handle_schedule(path)
                return
            if path.startswith("/api/monitor/"):
                self._handle_monitor(path)
                return
            self._send_json({"ok": False, "error": "未知接口"}, 404)
        except WebError as exc:
            self._send_json({"ok": False, "error": str(exc)}, exc.status)
        except Exception as exc:
            logger.exception("web POST 处理异常")
            self._send_json({"ok": False, "error": f"服务器错误：{exc}"}, 500)

    # ---------------- 各接口实现 ----------------
    def _handle_login(self):
        data = self._read_json()
        password = str(data.get("password", "") or "")
        token = self.service.verify_login(password, self._client_ip())
        cookie = (
            f"{SESSION_COOKIE}={token}; Path=/; Max-Age={SESSION_TTL_SEC}; "
            "HttpOnly; SameSite=Lax")
        self._send_json({"ok": True}, cookie=cookie)

    def _handle_upload(self, query):
        filename = (query.get("filename") or ["upload.xlsx"])[0]
        kind = (query.get("kind") or ["data"])[0]
        body = self._read_body(limit=MAX_BODY_BYTES)
        if not body:
            raise WebError("上传内容为空")
        upload_id = self.service.save_upload(filename, body, kind=kind)
        if kind == "attachment":
            self._send_json({
                "ok": True, "upload_id": upload_id,
                "filename": os.path.basename(filename),
            })
            return
        info = self.service.get_upload(upload_id) or {}
        names = self.service.read_sheet_names(info.get("path", ""))
        sheet = names[0] if names else ""
        grid = self.service.read_grid(info.get("path", ""), sheet)
        preview = self.service.build_preview(grid, 1)
        self._send_json({
            "ok": True,
            "upload_id": upload_id,
            "filename": info.get("name", ""),
            "sheets": names,
            "sheet": sheet,
            **preview,
        })

    def _handle_preview(self):
        data = self._read_json()
        upload_id = str(data.get("upload_id", "") or "")
        info = self.service.get_upload(upload_id)
        if info is None:
            raise WebError("上传文件已过期，请重新上传")
        sheet = str(data.get("sheet", "") or "")
        header_row = data.get("header_row", 1)
        grid = self.service.read_grid(info["path"], sheet)
        preview = self.service.build_preview(grid, header_row)
        sheets = self.service.read_sheet_names(info["path"])
        self._send_json({
            "ok": True, "upload_id": upload_id,
            "filename": info.get("name", ""),
            "sheets": sheets,
            "sheet": sheet or (sheets[0] if sheets else ""),
            **preview,
        })

    def _handle_names(self):
        """按人名所在列返回去重后的全部人员名（供网页勾选接收人）。"""
        data = self._read_json()
        upload_id = str(data.get("upload_id", "") or "")
        info = self.service.get_upload(upload_id)
        if info is None:
            raise WebError("上传文件已过期，请重新上传")
        sheet = str(data.get("sheet", "") or "")
        header_row = data.get("header_row", 1)
        name_column = str(data.get("name_column", "") or "").strip()
        if not name_column:
            raise WebError("请先选择人名所在列")
        try:
            header_row = max(1, int(header_row))
        except (TypeError, ValueError):
            header_row = 1
        grid = self.service.read_grid(info["path"], sheet)
        from .table_processor import TableProcessor
        try:
            processor = TableProcessor(grid, header_row=header_row)
            names = [str(x) for x in processor.get_all_persons(name_column, [])]
        except Exception as exc:
            raise WebError(f"读取人员列表失败：{exc}")
        self._send_json({"ok": True, "names": names})

    def _handle_cloud_preview(self):
        data = self._read_json()
        link = str(data.get("link", "") or "").strip()
        if not link:
            raise WebError("请填写金山在线表格链接")
        upload_id = self._download_cloud(link)
        info = self.service.get_upload(upload_id) or {}
        names = self.service.read_sheet_names(info["path"])
        sheet = names[0] if names else ""
        grid = self.service.read_grid(info["path"], sheet)
        preview = self.service.build_preview(grid, 1)
        self._send_json({
            "ok": True, "upload_id": upload_id,
            "filename": info.get("name", ""),
            "sheets": names, "sheet": sheet, **preview,
        })

    def _download_cloud(self, link: str) -> str:
        from .wps_cloud import (
            WpsAuthExpired, WpsCloudClient, WpsCloudError, parse_file_input,
        )
        try:
            token = parse_file_input(link)
        except WpsCloudError as exc:
            raise WebError(str(exc))
        temp_path = os.path.join(
            self.service._upload_dir,
            f"cloud_{secrets.token_hex(8)}.xlsx")
        client = WpsCloudClient(log_fn=lambda m: logger.info("[云文档] %s", m))
        try:
            client.download_to(token, temp_path)
        except WpsAuthExpired as exc:
            raise WebError(f"金山登录凭证已失效，请在电脑端重新配置：{exc}")
        except WpsCloudError as exc:
            raise WebError(f"云文档读取失败：{exc}")
        if not os.path.isfile(temp_path):
            raise WebError("云文档下载失败")
        upload_id = secrets.token_hex(12)
        with self.service._upload_lock:
            self.service._uploads[upload_id] = {
                "path": temp_path,
                "name": os.path.basename(temp_path),
                "kind": "cloud",
                "created": time.time(),
            }
            self.service._cleanup_uploads_locked()
        return upload_id

    def _handle_send(self):
        data = self._read_json()
        source = data.get("source") or {}
        stype = str(source.get("type", "upload") or "upload")
        upload_id = str(source.get("upload_id", "") or "")
        if stype == "cloud":
            link = str(source.get("link", "") or "").strip()
            if not link:
                raise WebError("请填写金山在线表格链接")
            upload_id = self._download_cloud(link)
        info = self.service.get_upload(upload_id)
        if info is None:
            raise WebError("数据文件已过期，请重新上传或预览")
        sheet = str(data.get("sheet", "") or "")
        header_row = data.get("header_row", 1)
        grid = self.service.read_grid(info["path"], sheet)

        payload = dict(data)
        payload["source"] = {"type": "upload", "upload_id": upload_id}
        payload["sheet"] = sheet or (self.service.read_sheet_names(info["path"]) or [""])[0]
        payload["header_row"] = header_row
        # 附件上传 ID → 本地路径（仅接受白名单类型）
        attachment_id = str(data.get("attachment_upload_id", "") or "")
        if attachment_id:
            att = self.service.get_upload(attachment_id)
            if att is not None:
                payload["attachment_path"] = att["path"]

        result = self.service.bridge.call(
            lambda: self.service.actions.send(payload, grid, info["path"]),
            timeout=30)
        if isinstance(result, dict):
            result.setdefault("ok", True)
        else:
            result = {"ok": True}
        self._send_json(result)

    def _handle_schedule(self, path):
        parts = [p for p in path.strip("/").split("/") if p]
        # /api/schedule/<id>/<action>
        if len(parts) != 4:
            self._send_json({"ok": False, "error": "接口不存在"}, 404)
            return
        task_id, action = parts[2], parts[3]
        if action == "run":
            data = self.service.bridge.call(
                lambda: self.service.actions.run_schedule(task_id), timeout=15)
            self._send_json(data)
            return
        if action == "enabled":
            body = self._read_json()
            enabled = bool(body.get("enabled", False))
            data = self.service.bridge.call(
                lambda: self.service.actions.set_schedule_enabled(task_id, enabled),
                timeout=10)
            data["ok"] = data.get("ok", True)
            self._send_json(data)
            return
        self._send_json({"ok": False, "error": "未知操作"}, 404)

    def _handle_monitor(self, path):
        parts = [p for p in path.strip("/").split("/") if p]
        if len(parts) != 4:
            self._send_json({"ok": False, "error": "接口不存在"}, 404)
            return
        task_id, action = parts[2], parts[3]
        if action == "run":
            data = self.service.bridge.call(
                lambda: self.service.actions.run_monitor(task_id), timeout=15)
            self._send_json(data)
            return
        if action == "enabled":
            body = self._read_json()
            enabled = bool(body.get("enabled", False))
            data = self.service.bridge.call(
                lambda: self.service.actions.set_monitor_enabled(task_id, enabled),
                timeout=10)
            data["ok"] = data.get("ok", True)
            self._send_json(data)
            return
        self._send_json({"ok": False, "error": "未知操作"}, 404)
