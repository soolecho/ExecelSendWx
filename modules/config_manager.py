import json
import os
from pathlib import Path
import tempfile
from datetime import datetime


class ConfigError(ValueError):
    pass


class ConfigManager:
    PROFILE_VERSION = 1
    MAX_RECENT = 10
    VALID_SEND_MODES = {"text", "image", "image_text", "custom"}

    def __init__(self, base_dir=None):
        self.base_dir = self._create_base_dir(base_dir)
        self.profile_dir = self.base_dir / "profiles"
        self.profile_dir.mkdir(parents=True, exist_ok=True)
        self.recent_file = self.base_dir / "recent_profiles.json"
        # 应用级全局设置（跨所有 profile）：拟人节流 + 全量默认映射等
        self.global_settings_file = self.base_dir / "global_settings.json"

    @staticmethod
    def _create_base_dir(base_dir):
        if base_dir:
            candidates = [Path(base_dir)]
        else:
            candidates = []
            local_app_data = os.environ.get("LOCALAPPDATA")
            if local_app_data:
                candidates.append(Path(local_app_data) / "ExcelSendWx")
            candidates.append(Path(tempfile.gettempdir()) / "ExcelSendWx")

        for candidate in candidates:
            try:
                candidate.mkdir(parents=True, exist_ok=True)
                return candidate
            except OSError:
                continue
        raise ConfigError("无法创建配置目录")

    @staticmethod
    def _absolute_path(path):
        return str(Path(path).expanduser().resolve(strict=False))

    @staticmethod
    def _write_json(path, data):
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_path = tempfile.mkstemp(
            prefix=f".{path.name}.",
            suffix=".tmp",
            dir=path.parent,
        )
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as file:
                json.dump(data, file, ensure_ascii=False, indent=2)
                file.write("\n")
            os.replace(temp_path, path)
        except Exception:
            try:
                os.remove(temp_path)
            except OSError:
                pass
            raise

    @classmethod
    def normalize_profile(cls, data):
        if not isinstance(data, dict):
            raise ConfigError("配置内容必须是 JSON 对象")

        version = data.get("version", cls.PROFILE_VERSION)
        if version != cls.PROFILE_VERSION:
            raise ConfigError(f"不支持的配置版本: {version}")

        excel = data.get("excel")
        send = data.get("send")
        if not isinstance(excel, dict) or not isinstance(send, dict):
            raise ConfigError("配置缺少 excel 或 send 设置")

        excel_path = str(excel.get("path", "")).strip()
        name_column = str(excel.get("name_column", "")).strip()
        extract_columns = excel.get("extract_columns", [])

        # 数据来源：local 本地文件（旧配置无此字段）/ wps_cloud 金山在线文档
        source = str(excel.get("source", "local")).strip()
        if source not in ("local", "wps_cloud"):
            source = "local"
        cloud_file_id = str(excel.get("cloud_file_id", "")).strip()
        cloud_file_name = str(excel.get("cloud_file_name", "")).strip()
        if source == "wps_cloud" and not cloud_file_id:
            # 兼容：历史异常配置缺少 ID 时回退本地来源
            source = "local"
        if source == "local" and not excel_path:
            raise ConfigError("配置中没有 Excel 文件路径")
        if not name_column:
            raise ConfigError("配置中没有人员列")
        if not isinstance(extract_columns, list):
            raise ConfigError("提取列格式不正确")

        extract_columns = [
            str(column).strip()
            for column in extract_columns
            if str(column).strip()
        ]
        if not extract_columns:
            raise ConfigError("配置中没有提取列")

        # 表头所在行（1-based）：旧配置无此字段时默认首行
        try:
            header_row = int(excel.get("header_row", 1))
        except (TypeError, ValueError):
            header_row = 1
        header_row = max(1, header_row)

        send_mode = str(send.get("mode", "text")).strip()
        if send_mode not in cls.VALID_SEND_MODES:
            raise ConfigError(f"不支持的发送形式: {send_mode}")

        try:
            send_interval = float(send.get("interval", 0.5))
            chat_delay = float(send.get("chat_delay", 0.3))
        except (TypeError, ValueError) as exc:
            raise ConfigError("延迟设置必须是数字") from exc

        filter_conditions = cls._normalize_filter_conditions(
            data.get("filter_conditions", [])
        )

        recipient_mapping = cls._normalize_recipient_mapping(
            data.get("recipient_mapping")
        )

        return {
            "version": cls.PROFILE_VERSION,
            "saved_at": str(data.get("saved_at", "")),
            "excel": {
                "path": excel_path,
                "sheet": str(excel.get("sheet", "")).strip(),
                "header_row": header_row,
                "name_column": name_column,
                "extract_columns": extract_columns,
                "wechat_column": str(
                    excel.get("wechat_column", "")
                ).strip(),
                "source": source,
                "cloud_file_id": cloud_file_id,
                "cloud_file_name": cloud_file_name,
            },
            "send": {
                "manual_recipient": str(
                    send.get("manual_recipient", "")
                ).strip(),
                "mode": send_mode,
                "interval": max(0.0, min(30.0, send_interval)),
                "chat_delay": max(0.0, min(10.0, chat_delay)),
                "custom_message_enabled": bool(
                    send.get("custom_message_enabled", False)
                ),
                "custom_message": str(send.get("custom_message", "")),
                "send_order": cls._normalize_send_order(send.get("send_order")),
                "attachment": str(send.get("attachment", "")).strip(),
                # 加载该配置后自动全选筛选人员并直接发送（无需手动点发送）
                "auto_send": bool(send.get("auto_send", False)),
            },
            "filter_conditions": filter_conditions,
            "recipient_mapping": recipient_mapping,
        }

    @staticmethod
    def _normalize_send_order(raw_order):
        """规范化发送顺序：只保留合法 key，去重保序。"""
        valid = {"text", "image", "custom", "attachment"}
        if not isinstance(raw_order, list):
            return None
        seen = set()
        result = []
        for k in raw_order:
            k = str(k).strip()
            if k in valid and k not in seen:
                seen.add(k)
                result.append(k)
        return result or None

    @staticmethod
    def _normalize_filter_conditions(raw):
        """规范化筛选条件：旧配置无此字段时返回空列表，保证向后兼容。"""
        if not isinstance(raw, list):
            return []
        valid_ops = {
            "equals", "contains", "not_contains", "empty", "not_empty",
            "greater_than", "less_than", "equals_or_greater",
            "equals_or_less", "date_equals", "date_after", "date_before",
        }
        result = []
        for item in raw:
            if not isinstance(item, dict):
                continue
            col = str(item.get("column_name", "")).strip()
            op = str(item.get("operator", "")).strip()
            val = str(item.get("value", ""))
            if col and op in valid_ops:
                result.append(
                    {"column_name": col, "operator": op, "value": val}
                )
        return result

    @staticmethod
    def _normalize_recipient_mapping(raw):
        """规范化联系人映射：旧配置无此字段时返回禁用状态的空映射，保证向后兼容。

        结构：
            enabled: bool         是否启用映射
            default_recipient: str  未命中映射且无 wechat_column 时的兜底接收人
            mapping_file: str    关联的外部映射表文件路径（用于"打开映射表"按钮）
            mappings: list       [{source_value, recipients: [str, ...]}]
        """
        if not isinstance(raw, dict):
            return {
                "enabled": False,
                "default_recipient": "",
                "mapping_file": "",
                "mappings": [],
            }
        mappings = []
        # 按 source_value 去重（后者覆盖前者），与发送逻辑保持一致
        dedup = {}
        for item in (raw.get("mappings") or []):
            if not isinstance(item, dict):
                continue
            src = str(item.get("source_value", "")).strip()
            recips_raw = item.get("recipients") or []
            if not isinstance(recips_raw, list):
                continue
            recips = [
                str(r).strip()
                for r in recips_raw
                if str(r).strip()
            ]
            if src and recips:
                dedup[src] = recips
        for src, recips in dedup.items():
            mappings.append(
                {"source_value": src, "recipients": recips}
            )
        return {
            "enabled": bool(raw.get("enabled", False)),
            "default_recipient": str(
                raw.get("default_recipient", "") or ""
            ).strip(),
            "mapping_file": str(
                raw.get("mapping_file", "") or ""
            ).strip(),
            "mappings": mappings,
        }

    # ------------------------------------------------------------------
    # 应用级全局设置（global_settings.json）：
    #  - rhythm: 拟人节流开关/档位/自定义随机延迟
    #  - default_mapping: 全量默认映射（仅作具体配置无映射时的最后兜底）
    # ------------------------------------------------------------------
    @staticmethod
    def _normalize_global_default_mapping(raw):
        """全量默认映射规范化：同一来源多行时**合并**接收人（不覆盖删除），保序。

        与 profile 的 _normalize_recipient_mapping（后者覆盖前者）不同——
        全局映射是用户在「全量设置」手打的文本，多行同源应累积而非丢弃，
        避免"保存后重新打开条目消失"。
        """
        if not isinstance(raw, dict):
            raw = {}
        merged = {}
        for item in (raw.get("mappings") or []):
            if not isinstance(item, dict):
                continue
            src = str(item.get("source_value", "") or "").strip()
            recips_raw = item.get("recipients") or []
            if not isinstance(recips_raw, list):
                continue
            recips = [
                str(r).strip()
                for r in recips_raw
                if str(r).strip()
            ]
            if src and recips:
                merged.setdefault(src, []).extend(recips)
        mappings = [
            {"source_value": src, "recipients": recips}
            for src, recips in merged.items()
        ]
        return {
            "enabled": bool(raw.get("enabled", False)),
            "default_recipient": str(
                raw.get("default_recipient", "") or ""
            ).strip(),
            "mapping_file": "",
            "mappings": mappings,
        }

    @staticmethod
    def _normalize_global_settings(raw):
        """规范化全局设置，返回一个含默认值的完整 dict。"""
        if not isinstance(raw, dict):
            raw = {}
        rhythm = raw.get("rhythm") or {}
        if not isinstance(rhythm, dict):
            rhythm = {}
        profile = str(rhythm.get("profile", "off")).strip()
        if profile not in ("off", "fast", "natural", "calm"):
            profile = "off"
        try:
            min_delay = float(rhythm.get("min_delay", 0.0))
            max_delay = float(rhythm.get("max_delay", 0.0))
        except (TypeError, ValueError):
            min_delay = 0.0
            max_delay = 0.0
        min_delay = max(0.0, min(30.0, min_delay))
        max_delay = max(0.0, min(30.0, max_delay))
        if max_delay < min_delay:
            max_delay = min_delay

        default_mapping = ConfigManager._normalize_global_default_mapping(
            raw.get("default_mapping")
        )

        return {
            "_version": 1,
            "rhythm": {
                "enabled": bool(rhythm.get("enabled", False)),
                "profile": profile,
                "min_delay": min_delay,
                "max_delay": max_delay,
            },
            "default_mapping": default_mapping,
        }

    def load_global_settings(self):
        """读取全局设置；文件不存在或损坏时返回默认值（并确保文件存在）。"""
        try:
            with self.global_settings_file.open(
                "r", encoding="utf-8"
            ) as file:
                data = json.load(file)
        except (OSError, json.JSONDecodeError):
            data = {}
        settings = self._normalize_global_settings(data)
        return settings

    def save_global_settings(self, settings):
        """保存全局设置（先规范化再原子写）。"""
        normalized = self._normalize_global_settings(settings)
        try:
            self._write_json(self.global_settings_file, normalized)
        except OSError as exc:
            raise ConfigError(f"全局设置保存失败: {exc}") from exc
        return normalized

    def save_profile(self, path, data):
        profile = self.normalize_profile(data)
        profile["saved_at"] = datetime.now().isoformat(timespec="seconds")
        profile_path = Path(path)
        if profile_path.suffix.lower() != ".json":
            profile_path = profile_path.with_suffix(".json")

        try:
            self._write_json(profile_path, profile)
        except OSError as exc:
            raise ConfigError(f"配置保存失败: {exc}") from exc

        profile_path = self._absolute_path(profile_path)
        self.add_recent(profile_path)
        return profile_path

    def load_profile(self, path):
        profile_path = Path(path)
        try:
            with profile_path.open("r", encoding="utf-8") as file:
                data = json.load(file)
        except (OSError, json.JSONDecodeError) as exc:
            raise ConfigError(f"配置读取失败: {exc}") from exc

        profile = self.normalize_profile(data)
        excel_path = Path(
            os.path.expandvars(profile["excel"]["path"])
        ).expanduser()
        if not excel_path.is_absolute():
            excel_path = profile_path.parent / excel_path
        profile["excel"]["path"] = self._absolute_path(excel_path)
        return profile

    def get_recent_profiles(self):
        try:
            with self.recent_file.open("r", encoding="utf-8") as file:
                data = json.load(file)
        except (OSError, json.JSONDecodeError):
            return []

        paths = data.get("paths", []) if isinstance(data, dict) else []
        if not isinstance(paths, list):
            return []

        recent = []
        seen = set()
        for path in paths:
            absolute_path = self._absolute_path(str(path))
            key = os.path.normcase(absolute_path)
            if key in seen or not os.path.isfile(absolute_path):
                continue
            seen.add(key)
            recent.append(absolute_path)
            if len(recent) >= self.MAX_RECENT:
                break
        return recent

    def add_recent(self, path):
        absolute_path = self._absolute_path(path)
        key = os.path.normcase(absolute_path)
        recent = [
            item
            for item in self.get_recent_profiles()
            if os.path.normcase(item) != key
        ]
        recent.insert(0, absolute_path)
        try:
            self._write_json(
                self.recent_file,
                {"paths": recent[:self.MAX_RECENT]},
            )
        except OSError:
            pass

    def remove_recent(self, path):
        key = os.path.normcase(self._absolute_path(path))
        recent = [
            item
            for item in self.get_recent_profiles()
            if os.path.normcase(item) != key
        ]
        try:
            self._write_json(self.recent_file, {"paths": recent})
        except OSError:
            pass
