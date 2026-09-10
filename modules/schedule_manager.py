r"""
定时发送配置与调度管理。

职责：
1. 独立存储定时任务到 %LOCALAPPDATA%\ExcelSendWx\schedules
2. 提供新增/修改/删除/启用/禁用接口
3. 定时调度：每 10 秒轮询当前时刻，匹配"每天/指定周几 + HH:MM"
4. 防重复执行：按日期+任务ID+时间点去重
5. 触发后通过信号交给 GUI 使用 WeChatSender 发送（含3次重试和模糊匹配）
"""
import json
import os
import tempfile
import threading
from dataclasses import dataclass, field, asdict
from datetime import datetime, date
from pathlib import Path
from typing import Dict, List, Optional


SCHEDULE_VERSION = 1

# 周一到周日，对应 isoweekday() 1..7
WEEKDAY_NAMES = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
WEEKDAY_MAP = {idx: name for idx, name in enumerate(WEEKDAY_NAMES, start=1)}


def _schedules_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        base = Path(local_app_data) / "ExcelSendWx" / "schedules"
    else:
        base = Path(tempfile.gettempdir()) / "ExcelSendWx" / "schedules"
    base.mkdir(parents=True, exist_ok=True)
    return base


SCHEDULE_INDEX_FILE = _schedules_dir() / "schedules.json"


@dataclass
class ScheduleTask:
    id: str
    name: str
    enabled: bool = True
    repeat_mode: str = "daily"  # daily | weekly
    # repeat_mode=daily 时 days 可为空；weekly 时 days 为 [1..7]
    days: List[int] = field(default_factory=list)
    # 多个时间点，HH:MM 字符串
    times: List[str] = field(default_factory=list)
    # 逗号/换行分隔的好友或群名
    recipients: List[str] = field(default_factory=list)
    message: str = ""
    chat_delay: float = 0.3
    # 每个接收人发送完成后间隔
    send_interval: float = 0.5
    # 任务级默认城市：{{weather}} 不带参时用此；为空则 fallback 到全局配置
    default_city: str = ""
    # 任务级电脑锁定配置（覆盖全局开关）
    keep_unlocked: bool = False  # 执行期间阻止电脑自动锁定
    relock_after: bool = False   # 发送完成后自动锁定电脑
    # 发送完成后最小化微信窗口（任务级开关）
    minimize_after: bool = True
    # 附加文件（文件/图片绝对路径），随消息额外发送；空表示不发送附件
    attachment: str = ""
    # 附件发送时机：after=先发文字再发附件（默认）；before=先发附件再发文字
    attach_order: str = "after"
    # 已触发记录：{ "YYYY-MM-DD": ["08:30", ...] }
    fired_log: Dict[str, List[str]] = field(default_factory=dict)
    created_at: str = ""
    updated_at: str = ""

    def to_dict(self) -> Dict:
        data = asdict(self)
        data["version"] = SCHEDULE_VERSION
        return data

    @classmethod
    def from_dict(cls, data: Dict) -> "ScheduleTask":
        def safe_list(value, cast=str):
            if not isinstance(value, list):
                return []
            items = []
            for v in value:
                try:
                    items.append(cast(v))
                except (TypeError, ValueError):
                    continue
            return items

        def safe_float(value, default):
            try:
                v = float(value)
                return max(0.0, min(30.0, v))
            except (TypeError, ValueError):
                return default

        days = safe_list(data.get("days"), int)
        days = sorted({d for d in days if 1 <= d <= 7})

        times: List[str] = []
        for raw in safe_list(data.get("times"), str):
            t = _normalize_time(raw)
            if t and t not in times:
                times.append(t)

        recipients: List[str] = []
        for raw in safe_list(data.get("recipients"), str):
            r = str(raw).strip()
            if r and r not in recipients:
                recipients.append(r)

        return cls(
            id=str(data.get("id", "")).strip() or _new_task_id(),
            name=str(data.get("name", "未命名任务")).strip() or "未命名任务",
            enabled=bool(data.get("enabled", True)),
            repeat_mode=str(data.get("repeat_mode", "daily"))
            if data.get("repeat_mode") in ("daily", "weekly")
            else "daily",
            days=days,
            times=times,
            recipients=recipients,
            message=str(data.get("message", "")),
            chat_delay=safe_float(data.get("chat_delay"), 0.3),
            send_interval=safe_float(data.get("send_interval"), 0.5),
            default_city=str(data.get("default_city", "") or ""),
            keep_unlocked=bool(data.get("keep_unlocked", False)),
            relock_after=bool(data.get("relock_after", False)),
            minimize_after=bool(data.get("minimize_after", True)),
            attachment=str(data.get("attachment", "") or ""),
            attach_order=str(data.get("attach_order", "after") or "after"),
            fired_log=_normalize_fired_log(data.get("fired_log")),
            created_at=str(data.get("created_at", "")),
            updated_at=str(data.get("updated_at", "")),
        )


def _new_task_id() -> str:
    import uuid
    return uuid.uuid4().hex[:16]


def _normalize_time(value: str) -> Optional[str]:
    if not value:
        return None
    v = str(value).strip().replace("：", ":")
    if ":" not in v:
        return None
    hh, mm = v.split(":", 1)
    try:
        h = int(hh.strip())
        m = int(mm.strip())
        if not (0 <= h <= 23 and 0 <= m <= 59):
            return None
        return f"{h:02d}:{m:02d}"
    except ValueError:
        return None


def _normalize_fired_log(value) -> Dict[str, List[str]]:
    if not isinstance(value, dict):
        return {}
    result: Dict[str, List[str]] = {}
    for k, v in value.items():
        try:
            datetime.strptime(str(k), "%Y-%m-%d")
        except ValueError:
            continue
        if not isinstance(v, list):
            continue
        slots: List[str] = []
        for t in v:
            slot = _normalize_time(str(t))
            if slot and slot not in slots:
                slots.append(slot)
        if slots:
            result[str(k)] = slots
    return result


def _write_json(path: Path, data) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=path.parent
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fp:
            json.dump(data, fp, ensure_ascii=False, indent=2)
            fp.write("\n")
        os.replace(temp_path, path)
    except Exception:
        try:
            os.remove(temp_path)
        except OSError:
            pass
        raise


class ScheduleStore:
    """定时任务配置存储，原子写入，线程安全。"""

    def __init__(self, index_file: Optional[Path] = None):
        self.index_file = Path(index_file) if index_file else SCHEDULE_INDEX_FILE
        self._lock = threading.Lock()

    def load_all(self) -> List[ScheduleTask]:
        with self._lock:
            try:
                with self.index_file.open("r", encoding="utf-8") as fp:
                    raw = json.load(fp)
            except (OSError, json.JSONDecodeError):
                return []
            if not isinstance(raw, dict):
                return []
            items = raw.get("tasks")
            if not isinstance(items, list):
                return []
            tasks = [ScheduleTask.from_dict(item) for item in items]
            tasks.sort(key=lambda t: (not t.enabled, t.updated_at or t.created_at, t.id))
            return tasks

    def _save_all(self, tasks: List[ScheduleTask]) -> None:
        data = {
            "version": SCHEDULE_VERSION,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "tasks": [t.to_dict() for t in tasks],
        }
        _write_json(self.index_file, data)

    def save(self, task: ScheduleTask) -> ScheduleTask:
        with self._lock:
            if not task.id:
                task.id = _new_task_id()
            now = datetime.now().isoformat(timespec="seconds")
            if not task.created_at:
                task.created_at = now
            task.updated_at = now
            tasks = self._load_all_unsafe()
            replaced = False
            for i, existing in enumerate(tasks):
                if existing.id == task.id:
                    tasks[i] = task
                    replaced = True
                    break
            if not replaced:
                tasks.append(task)
            self._save_all(tasks)
            return task

    def delete(self, task_id: str) -> bool:
        with self._lock:
            tasks = self._load_all_unsafe()
            new_tasks = [t for t in tasks if t.id != task_id]
            if len(new_tasks) == len(tasks):
                return False
            self._save_all(new_tasks)
            return True

    def mark_fired(self, task_id: str, day: str, slot: str) -> None:
        slot = _normalize_time(slot)
        if not slot:
            return
        try:
            datetime.strptime(day, "%Y-%m-%d")
        except ValueError:
            return
        with self._lock:
            tasks = self._load_all_unsafe()
            changed = False
            for t in tasks:
                if t.id != task_id:
                    continue
                slots = set(t.fired_log.get(day, []))
                if slot in slots:
                    # 已存在，避免无谓重写 JSON 与 update_at 抖动
                    return
                slots.add(slot)
                t.fired_log[day] = sorted(slots)
                # 只保留最近 30 天，避免无限增长
                before = len(t.fired_log)
                self._prune_fired_log(t)
                pruned = before != len(t.fired_log)
                t.updated_at = datetime.now().isoformat(timespec="seconds")
                changed = True
                # 记录 firing 但不改变其它字段
                _ = pruned
                break
            if changed:
                self._save_all(tasks)

    @staticmethod
    def _prune_fired_log(task: ScheduleTask) -> None:
        if not task.fired_log:
            return
        today = date.today()
        keep: Dict[str, List[str]] = {}
        for day, slots in task.fired_log.items():
            try:
                d = datetime.strptime(day, "%Y-%m-%d").date()
            except ValueError:
                continue
            if (today - d).days <= 30:
                keep[day] = slots
        task.fired_log = keep

    def _load_all_unsafe(self) -> List[ScheduleTask]:
        try:
            with self.index_file.open("r", encoding="utf-8") as fp:
                raw = json.load(fp)
        except (OSError, json.JSONDecodeError):
            return []
        if not isinstance(raw, dict):
            return []
        items = raw.get("tasks")
        if not isinstance(items, list):
            return []
        return [ScheduleTask.from_dict(item) for item in items]


class ScheduleDispatcher:
    """
    轮询调度器。

    不直接发送消息，只负责：
      - 判断“启用”、“今天/当前周几是否匹配”
      - 判断“当前时间是否落在某个 HH:MM 的 ±20 秒内”
      - 按日去重
    命中时通过回调列表通知外部执行发送。
    """

    def __init__(self, store: ScheduleStore):
        self.store = store
        self._handlers: List = []
        self._log_handlers: List = []
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._lock = threading.Lock()

    def add_handler(self, handler) -> None:
        """handler(task: ScheduleTask, slot: str) -> None"""
        with self._lock:
            if handler not in self._handlers:
                self._handlers.append(handler)

    def remove_handler(self, handler) -> None:
        with self._lock:
            if handler in self._handlers:
                self._handlers.remove(handler)

    def add_log_handler(self, handler) -> None:
        with self._lock:
            if handler not in self._log_handlers:
                self._log_handlers.append(handler)

    def _log(self, msg: str) -> None:
        with self._lock:
            handlers = list(self._log_handlers)
        for h in handlers:
            try:
                h(msg)
            except Exception:
                pass

    def start(self) -> None:
        with self._lock:
            if self._thread and self._thread.is_alive():
                return
            self._stop_event.clear()
            self._thread = threading.Thread(
                target=self._run, name="ScheduleDispatcher", daemon=True
            )
            self._thread.start()
        self._log("[定时] 调度器已启动")

    def stop(self) -> None:
        self._stop_event.set()
        t = self._thread
        if t and t.is_alive():
            t.join(timeout=2.0)
        self._log("[定时] 调度器已停止")

    def _run(self) -> None:
        # 启动先立即检查一次，避免整点错过
        self._tick()
        while not self._stop_event.is_set():
            # 每 10 秒检查一次；时间窗口为 ±20 秒，足够覆盖
            if self._stop_event.wait(10):
                break
            try:
                self._tick()
            except Exception as exc:  # pragma: no cover - 防御性
                self._log(f"[定时] 调度循环异常: {exc}")

    def _tick(self) -> None:
        now = datetime.now()
        today = now.strftime("%Y-%m-%d")
        current_hm = now.strftime("%H:%M")
        tasks = self.store.load_all()

        for task in tasks:
            if not task.enabled:
                continue
            if not self._match_today(task, now):
                continue
            fired_today = set(task.fired_log.get(today, []))
            for slot in task.times:
                if slot in fired_today:
                    continue
                if not self._in_time_window(now, slot):
                    continue
                self.store.mark_fired(task.id, today, slot)
                self._log(
                    f"[定时] 触发任务「{task.name}」时间点 {slot}，"
                    f"接收人数 {len(task.recipients)}"
                )
                with self._lock:
                    handlers = list(self._handlers)
                for h in handlers:
                    try:
                        h(task, slot)
                    except Exception as exc:
                        self._log(f"[定时] 调度回调异常: {exc}")
        # 说明：普通扫描无命中时不输出日志，避免每 10 秒刷一次导致日志与 GUI 事件循环压力。
        # 只保留 fired/start/stop/error 这几类关键信息；如需要调试请手动看这里加 debug 输出。

    @staticmethod
    def _match_today(task: ScheduleTask, now: datetime) -> bool:
        if task.repeat_mode == "daily":
            return True
        if task.repeat_mode == "weekly":
            return now.isoweekday() in set(task.days)
        return False

    @staticmethod
    def _in_time_window(now: datetime, slot: str) -> bool:
        try:
            hour_str, minute_str = slot.split(":", 1)
            slot_dt = now.replace(
                hour=int(hour_str),
                minute=int(minute_str),
                second=0,
                microsecond=0,
            )
        except (ValueError, AttributeError):
            return False
        delta = abs((now - slot_dt).total_seconds())
        return delta <= 20.0
