"""监控配置模型与持久化。

监控场景：企业微信客户端「聊天文件自动下载」会把群文件自动保存到本地
固定目录。本模块负责新建/编辑/删除监控任务，并持久化到
%LOCALAPPDATA%\\ExcelSendWx\\monitors\\monitors.json。

每个监控任务：监控一个本地目录 → 按规则筛选数据 → 可选与上次对比找增量
→ 有新内容才推送（文字/图片/生成表格文件）到个人微信（复用 wxauto）。

时间窗语义与 ScheduleTask 保持一致（daily/weekly/once + HH:MM 时间点），
空 times 表示 24 小时轮询；poll_interval_min 控制扫描周期。
"""
import json
import os
import tempfile
import threading
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional


MONITOR_VERSION = 2


@dataclass
class CleanRule:
    """单条数据清洗/校验规则。

    加工顺序：去空(str.replace(' ')) -> 去非数字(keep_digits_only) ->
    取前 N 位(take_first_n>0) -> 前缀校验(require_prefix 非空时 startswith)。
    on_fail="drop" 整行剔除；on_fail="keep" 保留清洗后的值但不本行剔除外。
    """
    column: str = ""
    strip_space: bool = True
    keep_digits_only: bool = False
    take_first_n: int = 0        # 0=不截取
    require_prefix: str = ""
    on_fail: str = "drop"        # "drop" | "keep"

    def to_dict(self) -> Dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict) -> "CleanRule":
        def _s(key, default=""):
            v = data.get(key)
            return str(v) if v is not None else default

        def _b(key, default):
            v = data.get(key)
            return bool(v) if isinstance(v, bool) else default

        def _i(key, default):
            try:
                return int(data.get(key, default))
            except (TypeError, ValueError):
                return default

        on_fail = _s("on_fail", "drop")
        if on_fail not in ("drop", "keep"):
            on_fail = "drop"
        return cls(
            column=_s("column"),
            strip_space=_b("strip_space", True),
            keep_digits_only=_b("keep_digits_only", False),
            take_first_n=max(0, _i("take_first_n", 0)),
            require_prefix=_s("require_prefix"),
            on_fail=on_fail,
        )


def _monitors_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        base = Path(local_app_data) / "ExcelSendWx" / "monitors"
    else:
        base = Path(tempfile.gettempdir()) / "ExcelSendWx" / "monitors"
    base.mkdir(parents=True, exist_ok=True)
    return base


MONITOR_INDEX_FILE = _monitors_dir() / "monitors.json"


def _new_task_id() -> str:
    return datetime.now().strftime("mon_%Y%m%d%H%M%S%f")


def _split_list(value) -> List[str]:
    """逗号/分号/换行分隔多值 → 去空去重列表。"""
    if isinstance(value, list):
        items = value
    else:
        items = str(value or "").split()
    result = []
    seen = set()
    for it in items:
        for part in str(it).replace("；", ";").replace("，", ",").replace("\n", ",").split(","):
            part = part.strip()
            if part and part not in seen:
                seen.add(part)
                result.append(part)
    return result


@dataclass
class MonitorTask:
    id: str
    name: str
    enabled: bool = True
    # 文件来源
    watch_path: str = ""
    file_pattern: str = "*.xlsx;*.xls"
    # 筛选规则：sheet_name 空=第一个 sheet；filter_column 空=不筛选(全表)
    sheet_name: str = ""
    filter_column: str = ""
    filter_values: List[str] = field(default_factory=list)
    # 对比增量（可选项）：compare_enabled=False 则不对比，检测到新文件即推送
    compare_enabled: bool = True
    # 开启对比时使用：baseline[file_key] -> List[row_key] 已处理行键集合
    baseline: Dict[str, List[str]] = field(default_factory=dict)
    # 关闭对比时使用：seen_files[文件名] -> 指纹（size/mtime），判定"新文件"
    seen_files: Dict[str, str] = field(default_factory=dict)
    # 推送选项（可多选）
    send_text: bool = True
    send_image: bool = False
    send_file: bool = False
    # 推送目标（个人微信名）
    recipients: List[str] = field(default_factory=list)
    # 时间窗 + 轮询周期：times 空=全天；repeat_mode/days/run_dates 同 ScheduleTask
    repeat_mode: str = "daily"
    days: List[int] = field(default_factory=list)
    run_dates: List[str] = field(default_factory=list)
    times: List[str] = field(default_factory=list)
    poll_interval_min: int = 5
    # 发送参数
    chat_delay: float = 0.3
    send_interval: float = 0.5
    minimize_after: bool = True
    keep_unlocked: bool = False
    relock_after: bool = False
    # 数据清洗/校验规则（阶段B）
    clean_rules: List[CleanRule] = field(default_factory=list)
    # 本地台账追加（阶段B）
    ledger_enabled: bool = False
    ledger_path: str = ""
    ledger_sheet: str = ""
    ledger_remark_col: str = ""
    ledger_remark_text: str = ""
    ledger_exclude_extra: bool = False
    created_at: str = ""
    updated_at: str = ""

    # ------- 序列化 -------
    def to_dict(self) -> Dict:
        data = asdict(self)
        data["version"] = MONITOR_VERSION
        return data

    @classmethod
    def from_dict(cls, data: Dict) -> "MonitorTask":
        def _s(key, default=""):
            v = data.get(key)
            return str(v) if v is not None else default

        def _b(key, default):
            v = data.get(key)
            return bool(v) if isinstance(v, bool) else default

        def _f(key, default):
            try:
                return float(data.get(key, default))
            except (TypeError, ValueError):
                return default

        def _i(key, default):
            try:
                return int(data.get(key, default))
            except (TypeError, ValueError):
                return default

        days = [d for d in data.get("days", []) if isinstance(d, int)] if isinstance(
            data.get("days"), list
        ) else []
        days = sorted(d for d in set(days) if 1 <= d <= 7)
        run_dates = sorted(
            d for d in dict.fromkeys(str(x).strip() for x in data.get("run_dates", [])) if d
        )
        times = _split_list(data.get("times"))
        base = data.get("baseline")
        seen = data.get("seen_files")
        task = cls(
            id=_s("id", _new_task_id()),
            name=_s("name", "未命名监控"),
            enabled=_b("enabled", True),
            watch_path=_s("watch_path"),
            file_pattern=_s("file_pattern", "*.xlsx;*.xls"),
            sheet_name=_s("sheet_name"),
            filter_column=_s("filter_column"),
            filter_values=_split_list(data.get("filter_values")),
            compare_enabled=_b("compare_enabled", True),
            baseline={
                str(k): list(v) for k, v in base.items()
            } if isinstance(base, dict) else {},
            seen_files={
                str(k): str(v) for k, v in seen.items()
            } if isinstance(seen, dict) else {},
            send_text=_b("send_text", True),
            send_image=_b("send_image", False),
            send_file=_b("send_file", False),
            recipients=_split_list(data.get("recipients")),
            repeat_mode=_s("repeat_mode", "daily") or "daily",
            days=days,
            run_dates=run_dates,
            times=times,
            poll_interval_min=max(1, _i("poll_interval_min", 5)),
            chat_delay=_f("chat_delay", 0.3),
            send_interval=_f("send_interval", 0.5),
            minimize_after=_b("minimize_after", True),
            keep_unlocked=_b("keep_unlocked", False),
            relock_after=_b("relock_after", False),
            clean_rules=[
                CleanRule.from_dict(c) for c in data.get("clean_rules", [])
                if isinstance(c, dict)
            ],
            ledger_enabled=_b("ledger_enabled", False),
            ledger_path=_s("ledger_path"),
            ledger_sheet=_s("ledger_sheet"),
            ledger_remark_col=_s("ledger_remark_col"),
            ledger_remark_text=_s("ledger_remark_text"),
            ledger_exclude_extra=_b("ledger_exclude_extra", False),
            created_at=_s("created_at"),
            updated_at=_s("updated_at"),
        )
        return task


def _now_in_window(task: MonitorTask, now: datetime) -> bool:
    """判断当前是否为任务的活动监控日。

    daily=每天；weekly=指定的周几；once=指定的执行日期之一。
    监控采用轮询（poll_interval_min 节流），times 字段暂不参与日期判定，
    保留仅供将来按时间片收窄使用。
    """
    if task.repeat_mode == "weekly":
        return now.isoweekday() in set(task.days)
    if task.repeat_mode == "once":
        return now.strftime("%Y-%m-%d") in set(task.run_dates)
    return True  # daily


class MonitorManager:
    """监控任务持久化管理（镜像 ScheduleManager 风格）。"""

    def __init__(self, index_file: Optional[Path] = None):
        self.index_file = Path(index_file) if index_file else MONITOR_INDEX_FILE
        self._lock = threading.Lock()

    def load_all(self) -> List[MonitorTask]:
        with self._lock:
            return self._load_all_unsafe()

    def _load_all_unsafe(self) -> List[MonitorTask]:
        try:
            with self.index_file.open("r", encoding="utf-8") as fp:
                raw = json.load(fp)
        except (OSError, json.JSONDecodeError):
            return []
        if not isinstance(raw, dict):
            return []
        items = raw.get("monitors")
        if not isinstance(items, list):
            return []
        tasks = [MonitorTask.from_dict(item) for item in items]
        tasks.sort(key=lambda t: (not t.enabled, t.updated_at or t.created_at, t.id))
        return tasks

    def _save_all(self, tasks: List[MonitorTask]) -> None:
        data = {
            "version": MONITOR_VERSION,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "monitors": [t.to_dict() for t in tasks],
        }
        try:
            with self.index_file.open("w", encoding="utf-8") as fp:
                json.dump(data, fp, ensure_ascii=False, indent=2)
        except OSError:
            pass

    def save(self, task: MonitorTask) -> MonitorTask:
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

    def get(self, task_id: str) -> Optional[MonitorTask]:
        for t in self._load_all_unsafe():
            if t.id == task_id:
                return t
        return None

    def delete(self, task_id: str) -> bool:
        with self._lock:
            tasks = self._load_all_unsafe()
            new_tasks = [t for t in tasks if t.id != task_id]
            if len(new_tasks) == len(tasks):
                return False
            self._save_all(new_tasks)
            return True

    def set_enabled(self, task_id: str, enabled: bool) -> bool:
        with self._lock:
            tasks = self._load_all_unsafe()
            changed = False
            for t in tasks:
                if t.id == task_id:
                    t.enabled = enabled
                    t.updated_at = datetime.now().isoformat(timespec="seconds")
                    changed = True
                    break
            if changed:
                self._save_all(tasks)
            return changed