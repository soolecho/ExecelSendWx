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


MONITOR_VERSION = 4


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
    # 文件名前缀过滤：非空时只处理文件名以该前缀开头的文件；留空=全量
    file_prefix: str = ""
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
    # 对比列：决定"新增"用哪些列（空 = 全部列）
    compare_columns: List[str] = field(default_factory=list)
    # 数据清洗/校验规则（每条规则针对一列，由"列设置"面板统一生成）
    clean_rules: List[CleanRule] = field(default_factory=list)
    # 发送提取列：只把这几列写入推送文字/图片/文件（空 = 全部列）
    extract_columns: List[str] = field(default_factory=list)
    # 推送文字自定义标题（如"光缆故障新增提醒"），未配置用默认
    text_title: str = ""
    # 推送文字是否附带逐行明细（勾选才显示 列:值；否则只显示标题+条数）
    text_detail: bool = True
    # 是否递归扫描监控目录的子文件夹（默认仅当前层，避免大目录过慢）
    include_subdir: bool = False
    # 每天运行时间段（HH:MM，起:止）。空=全天；支持跨夜如 22:00-07:00。
    active_start: str = ""
    active_end: str = ""
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
            file_prefix=_s("file_prefix"),
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
            compare_columns=_split_list(data.get("compare_columns")),
            extract_columns=_split_list(data.get("extract_columns")),
            text_title=_s("text_title"),
            text_detail=_b("text_detail", True),
            include_subdir=_b("include_subdir", False),
            active_start=_s("active_start"),
            active_end=_s("active_end"),
            created_at=_s("created_at"),
            updated_at=_s("updated_at"),
        )
        return task


def _now_in_window(task: MonitorTask, now: datetime) -> bool:
    """判断当前是否为任务的活动监控日。

    daily=每天；weekly=指定的周几；once=指定的执行日期之一。
    监控采用轮询（poll_interval_min 节流），times 字段暂不参与日期判定，
    保留仅供将来按时间片收窄使用。

    每天运行时间段（active_start~active_end，HH:MM）：空=全天；起止相同=
    全天；支持跨夜（如 22:00-07:00 表示晚上10点到次日早上7点之间活动）。
    """
    # 时段窗口判定（对所有 repeat_mode 生效）
    s = str(task.active_start or "").strip()
    e = str(task.active_end or "").strip()
    if s and e and s != e:
        cur = now.strftime("%H:%M")
        if s < e:
            if not (s <= cur <= e):
                return False
        else:  # 跨夜
            if not (cur >= s or cur <= e):
                return False

    if task.repeat_mode == "weekly":
        return now.isoweekday() in set(task.days)
    if task.repeat_mode == "once":
        return now.strftime("%Y-%m-%d") in set(task.run_dates)
    return True  # daily


def _profiles_dir() -> Path:
    """独立配置文件目录：monitors/profiles/<名称>.json，一个监控配置一个文件。"""
    base = _monitors_dir() / "profiles"
    base.mkdir(parents=True, exist_ok=True)
    return base


def _sanitize_name(name: str) -> str:
    name = "".join(c for c in (name or "")
                   if c not in '\\/:*?"<>|').strip()
    return name or "监控"


def _profile_path(name: str) -> Path:
    return _profiles_dir() / (_sanitize_name(name) + ".json")


class MonitorManager:
    """监控配置文件管理：每个监控配置 = profiles 目录下一个独立的 .json 文件。

    左侧列表 = 扫描 profiles 目录；保存/加载针对单个配置文件，便于导出、导入、
    切换运行不同的监控配置（符合"能运行不同的监控配置文件"诉求）。
    worker 的 load_all() 返回全部配置作为常驻运行池。
    旧版单索引 monitors.json 若存在且无独立配置时，自动迁移为独立文件。
    """

    def __init__(self, index_file: Optional[Path] = None):
        # index_file 仅作兼容参数保留；实际改用 profiles 目录
        self._lock = threading.Lock()
        self._migrate_legacy()

    # ---------------- 旧单索引自动迁移 ----------------
    def _migrate_legacy(self) -> None:
        try:
            with self._lock:
                legacy = MONITOR_INDEX_FILE
                if not legacy.exists():
                    return
                if list(_profiles_dir().glob("*.json")):
                    return
                with legacy.open("r", encoding="utf-8") as fp:
                    raw = json.load(fp)
                items = raw.get("monitors") if isinstance(raw, dict) else None
                if isinstance(items, list):
                    for item in items:
                        try:
                            self._save_file(MonitorTask.from_dict(item))
                        except Exception:
                            continue
                try:
                    legacy.unlink()
                except OSError:
                    pass
        except Exception:
            pass

    # ---------------- 列表 ----------------
    def load_all(self) -> List[MonitorTask]:
        with self._lock:
            tasks = []
            for p in sorted(
                _profiles_dir().glob("*.json"),
                key=lambda x: x.stat().st_mtime if x.exists() else 0,
            ):
                try:
                    with p.open("r", encoding="utf-8") as fp:
                        raw = json.load(fp)
                    tasks.append(MonitorTask.from_dict(raw))
                except Exception:
                    continue
            tasks.sort(key=lambda t: (not t.enabled, t.name.lower()))
            return tasks

    def list_profiles(self) -> List[dict]:
        with self._lock:
            out = []
            for p in sorted(_profiles_dir().glob("*.json")):
                d = {"path": str(p), "name": p.stem, "enabled": False, "updated_at": ""}
                try:
                    with p.open("r", encoding="utf-8") as fp:
                        raw = json.load(fp)
                    d["enabled"] = bool(raw.get("enabled", True))
                    d["updated_at"] = str(raw.get("updated_at", ""))
                except Exception:
                    pass
                out.append(d)
            return out

    # ---------------- 存取 ----------------
    def _save_file(self, task: MonitorTask) -> None:
        p = _profile_path(task.name)
        with p.open("w", encoding="utf-8") as fp:
            json.dump(task.to_dict(), fp, ensure_ascii=False, indent=2)

    def save(self, task: MonitorTask) -> MonitorTask:
        with self._lock:
            if not task.id:
                task.id = _new_task_id()
            now = datetime.now().isoformat(timespec="seconds")
            if not task.created_at:
                task.created_at = now
            task.updated_at = now
            # 改名或旧文件位置不同 → 删除旧 id 对应文件，避免残留
            new_name = _sanitize_name(task.name) + ".json"
            for p in _profiles_dir().glob("*.json"):
                if p.name == new_name:
                    continue
                try:
                    with p.open("r", encoding="utf-8") as fp:
                        raw = json.load(fp)
                except Exception:
                    continue
                if str(raw.get("id")) == task.id:
                    try:
                        p.unlink()
                    except OSError:
                        pass
            self._save_file(task)
            return task

    def get(self, task_id: str) -> Optional[MonitorTask]:
        for t in self.load_all():
            if t.id == task_id:
                return t
        return None

    def load_profile(self, name: str) -> Optional[MonitorTask]:
        p = _profile_path(name)
        if not p.exists():
            return None
        with self._lock:
            try:
                with p.open("r", encoding="utf-8") as fp:
                    raw = json.load(fp)
                return MonitorTask.from_dict(raw)
            except Exception:
                return None

    def delete(self, task_id: str) -> bool:
        with self._lock:
            for p in _profiles_dir().glob("*.json"):
                try:
                    with p.open("r", encoding="utf-8") as fp:
                        raw = json.load(fp)
                except Exception:
                    continue
                if str(raw.get("id")) == task_id:
                    try:
                        p.unlink()
                        return True
                    except OSError:
                        return False
            return False

    def set_enabled(self, task_id: str, enabled: bool) -> bool:
        with self._lock:
            for p in _profiles_dir().glob("*.json"):
                try:
                    with p.open("r", encoding="utf-8") as fp:
                        raw = json.load(fp)
                except Exception:
                    continue
                if str(raw.get("id")) == task_id:
                    raw["enabled"] = bool(enabled)
                    raw["updated_at"] = datetime.now().isoformat(timespec="seconds")
                    try:
                        with p.open("w", encoding="utf-8") as fp:
                            json.dump(raw, fp, ensure_ascii=False, indent=2)
                        return True
                    except OSError:
                        return False
            return False

    def save_profile_as(self, task: MonitorTask) -> str:
        """另存为：以当前 name 落成独立配置文件，返回文件路径。"""
        with self._lock:
            if not task.id:
                task.id = _new_task_id()
            if not task.created_at:
                task.created_at = datetime.now().isoformat(timespec="seconds")
            task.updated_at = task.created_at
            self._save_file(task)
            return str(_profile_path(task.name))