# -*- coding: utf-8 -*-
"""配置文件（profile）→ 可发送任务列表 的无头准备逻辑。

供"加载配置后自动发送"和"定时/链式执行配置文件"共用：

    本地 xlsx/云端表格 → 读取全部 sheet（云端读完即删临时文件）
    → TableProcessor(表头行) → 按配置筛选人员
    → 逐人提取表格数据、确定微信接收人
    → 输出与 SendWorker 契约一致的 task dict 列表

本模块不依赖 Qt，可在任意 worker 线程调用。
"""

import os
import time

from modules.table_processor import TableProcessor, FilterCondition
from modules.wps_cloud import (
    WpsOAuthStore,
    WpsCloudClient,
    new_temp_xlsx_path,
    cleanup_temp_file,
    read_one_sheet,
    read_sheet_names,
)


# 联系人映射多值分隔符（与 gui._split_rm_values 保持一致）
_RM_SEPARATORS = ("/", ";", "、", "，", ",")


def _split_rm_values(raw):
    """按 / ; 、 ， , 拆分多值，去空白和空串。"""
    if not raw:
        return []
    text = str(raw).strip()
    if not text:
        return []
    for sep in _RM_SEPARATORS:
        text = text.replace(sep, "\n")
    return [s.strip() for s in text.splitlines() if s.strip()]


def _load_mappings_from_file(path, log):
    """从映射表文件读取两列 → list[{source_value, recipients}]。

    文件两列：A=筛选值（可多值分隔），B=接收人（可多值分隔）。
    读取失败返回 None，调用方用配置里的快照 mappings 兜底。
    """
    try:
        names = read_sheet_names(path, log_fn=log)
        if not names:
            log(f"⚠ 映射表无工作表: {path}")
            return None
        rows = read_one_sheet(path, names[0], log_fn=log)
    except Exception as exc:
        log(f"⚠ 读取映射表失败，使用配置快照: {exc}")
        return None
    if not rows:
        return None
    merged = {}
    for row in rows:
        if not row or len(row) < 2:
            continue
        srcs = _split_rm_values(row[0])
        recips = _split_rm_values(row[1])
        if not srcs or not recips:
            continue
        for s in srcs:
            merged[s] = list(recips)
    return [{"source_value": s, "recipients": r} for s, r in merged.items()]


class ProfilePrepareError(Exception):
    """配置无法执行（文件缺失/列不匹配/数据为空等）。"""


def _render_weather(custom_msg, log):
    """自定义消息中的天气/日期占位符渲染（与定时消息同一套渲染器）。"""
    if not custom_msg:
        return custom_msg
    try:
        from modules import weather_fetcher
        if not weather_fetcher.find_placeholders(custom_msg):
            return custom_msg
        cfg = weather_fetcher.load_weather_config()
        rendered = weather_fetcher.render_message(
            custom_msg,
            task_default_city="",
            api_key=cfg.get("api_key", ""),
            base_url=cfg.get("base_url", ""),
            global_default_city=cfg.get("default_city", ""),
        )
        return rendered
    except Exception as exc:
        log(f"占位符渲染异常（保留原文）: {exc}")
        return custom_msg


def _resolve_send_order(send):
    """与 GUI _apply_pending_config 同一套 send_order 兜底/裁剪规则。"""
    custom_enabled = bool(send.get("custom_message_enabled", False))
    custom_msg = str(send.get("custom_message", "") or "")
    attachment = str(send.get("attachment", "") or "").strip()

    raw_order = send.get("send_order")
    if isinstance(raw_order, list) and raw_order:
        order = [
            str(k).strip()
            for k in raw_order
            if str(k).strip() in ("text", "image", "custom", "attachment")
        ]
        if custom_enabled and custom_msg and "custom" not in order:
            order.append("custom")
        if attachment and "attachment" not in order:
            order.append("attachment")
    else:
        mode = str(send.get("mode", "text"))
        if mode == "image":
            order = ["image"]
        elif mode == "image_text":
            order = ["image", "text"]
        else:
            order = ["text"]
        if custom_enabled and custom_msg:
            order.append("custom")
        if attachment:
            order.append("attachment")

    if not (custom_enabled and custom_msg):
        order = [k for k in order if k != "custom"]
    if not attachment or not os.path.exists(attachment):
        order = [k for k in order if k != "attachment"]
    # 去重保序
    result = []
    for k in order:
        if k not in result:
            result.append(k)
    return result


def prepare_profile(profile, log_fn=None):
    """把一个规范化后的 profile 转成可直接交给 send_executor 的载荷。

    返回 dict：
      tasks / send_order / attachment / send_interval / chat_delay / title
    """
    log = log_fn or (lambda msg: None)
    excel = profile["excel"]
    send = profile["send"]

    source = excel.get("source", "local")
    temp_path = None
    try:
        # 1) 读取工作簿：配置链/auto_send 只需要配置指定的那一个 sheet。
        #    大工作簿（实测 37MB/13 sheet）全量解析近 100 秒并吃满 GIL，
        #    会把界面拖成“未响应”；定向读取仅需数秒。
        saved_sheet = (excel.get("sheet") or "").strip()
        if source == "wps_cloud":
            file_token = (excel.get("cloud_file_id") or "").strip()
            if not file_token:
                raise ProfilePrepareError("云文档配置缺少文件 ID")
            title = excel.get("cloud_file_name") or file_token
            log(f"下载云文档: {title}")
            store = WpsOAuthStore()
            client = WpsCloudClient(store, log_fn=log)
            temp_path = new_temp_xlsx_path()
            t0 = time.time()
            client.download_to(file_token, temp_path)
            log(f"云文档下载完成（{time.time() - t0:.0f}秒），正在解析工作表…")
            workbook_path = temp_path
        else:
            workbook_path = os.path.expandvars(excel.get("path", ""))
            if not os.path.isfile(workbook_path):
                raise ProfilePrepareError(f"Excel 文件不存在: {workbook_path}")
            title = os.path.splitext(os.path.basename(workbook_path))[0]

        t0 = time.time()
        rows = None
        sheet_name = saved_sheet
        if saved_sheet:
            rows = read_one_sheet(workbook_path, saved_sheet, log_fn=log)
            if rows is None:
                # 配置里的 sheet 名失效：列名单后取第一个，并提示
                names = read_sheet_names(workbook_path, log_fn=log)
                if not names:
                    raise ProfilePrepareError("工作簿中没有工作表")
                sheet_name = names[0]
                log(f"配置的 Sheet“{saved_sheet}”不存在，使用“{sheet_name}”")
                rows = read_one_sheet(workbook_path, sheet_name, log_fn=log)
        else:
            # 未指定 sheet：只取第一个（避免全量读取大工作簿）
            names = read_sheet_names(workbook_path, log_fn=log)
            if not names:
                raise ProfilePrepareError("工作簿中没有工作表")
            sheet_name = names[0]
            rows = read_one_sheet(workbook_path, sheet_name, log_fn=log)
        log(f"工作表“{sheet_name}”解析完成（{time.time() - t0:.0f}秒）")
        if not rows:
            raise ProfilePrepareError(f"工作表 {sheet_name} 为空")

        # 2) 表头行 + 处理器
        header_row = max(1, int(excel.get("header_row", 1) or 1))
        processor = TableProcessor(rows, header_row=header_row)
        headers = set(processor.get_headers())

        name_column = (excel.get("name_column") or "").strip()
        if name_column not in headers:
            raise ProfilePrepareError(f"人员列“{name_column}”不存在")

        extract_columns = [
            c for c in (excel.get("extract_columns") or []) if c in headers
        ]
        missing_cols = [
            c for c in (excel.get("extract_columns") or []) if c not in headers
        ]
        if missing_cols:
            log(f"部分提取列不存在，已忽略: {', '.join(missing_cols)}")
        if not extract_columns:
            raise ProfilePrepareError("没有可用的提取列")

        # 3) 筛选条件
        conditions = []
        for cond in profile.get("filter_conditions", []) or []:
            col = str(cond.get("column_name", "")).strip()
            op = str(cond.get("operator", "")).strip()
            if col and op:
                conditions.append(FilterCondition(col, op, str(cond.get("value", ""))))

        persons = processor.get_all_persons(name_column, conditions)
        log(f"「{title}」筛选后共 {len(persons)} 人")
        if not persons:
            raise ProfilePrepareError("筛选结果为空，没有符合条件的人员")

        # 4) 微信接收人映射
        wechat_column = (excel.get("wechat_column") or "").strip()
        mapping = {}
        if wechat_column:
            if wechat_column in headers:
                mapping = processor.get_person_to_wechat_mapping(
                    name_column, wechat_column
                )
            else:
                log(f"微信列“{wechat_column}”不存在，忽略")
        manual_recipient = (send.get("manual_recipient") or "").strip()

        # 4.5) 联系人映射（recipient_mapping）：把"人员列"筛选出来的值
        #      （可能是组名"装维组"等）映射到实际的微信接收人列表。
        #      - 启用映射且命中 → 一个 person 可能展开为多个 task（一对多）
        #      - 启用映射但未命中且有兜底 → 用兜底接收人
        #      - 启用映射但未命中且无兜底 / 未启用映射 → 走原 wechat_column → manual_recipient → person 逻辑
        mapping_cfg = profile.get("recipient_mapping") or {}
        mapping_enabled = bool(mapping_cfg.get("enabled", False))
        value_to_recipients = {}
        if mapping_enabled:
            mappings = list(mapping_cfg.get("mappings") or [])
            # 关联了映射表文件 → 发送前自动读取最新内容，
            # 用户改了文件不用手动"导入表格"再保存配置
            mapping_file = str(mapping_cfg.get("mapping_file", "") or "").strip()
            if mapping_file and os.path.isfile(mapping_file):
                fresh = _load_mappings_from_file(mapping_file, log)
                if fresh:
                    # 覆盖式合并：文件覆盖同名 key，配置内嵌映射中文件没有的保留
                    file_keys = {
                        str(m.get("source_value", "") or "").strip()
                        for m in fresh
                    }
                    mappings = fresh + [
                        m for m in mappings
                        if str(m.get("source_value", "") or "").strip() not in file_keys
                    ]
                    log(f"已从关联文件自动加载 {len(fresh)} 条映射: {mapping_file}")
            for m in mappings:
                src = str(m.get("source_value", "") or "").strip()
                recips_raw = m.get("recipients") or []
                if not isinstance(recips_raw, list):
                    continue
                recips = [
                    str(r).strip()
                    for r in recips_raw
                    if str(r).strip()
                ]
                if src and recips:
                    value_to_recipients[src] = recips
            log(
                f"联系人映射已启用：{len(value_to_recipients)} 条映射，"
                f"兜底接收人={mapping_cfg.get('default_recipient', '') or '（无）'}"
            )
        default_recipient = str(
            mapping_cfg.get("default_recipient", "") or ""
        ).strip()

        # 全量默认映射兜底：profile 自身未启用映射（或未命中）时，
        # 回退到应用级「全量设置」的默认映射（最后兜底层）。
        if not mapping_enabled:
            from modules.config_manager import ConfigManager
            try:
                gs_cfg = ConfigManager().load_global_settings()
            except Exception:
                gs_cfg = {}
            g_dm = gs_cfg.get("default_mapping") or {}
            if g_dm.get("enabled"):
                g_v2r = {}
                for m in (g_dm.get("mappings") or []):
                    src = str(m.get("source_value", "") or "").strip()
                    recips = [
                        str(r).strip()
                        for r in (m.get("recipients") or [])
                        if str(r).strip()
                    ]
                    if src and recips:
                        g_v2r[src] = recips
                g_default = str(
                    g_dm.get("default_recipient", "") or ""
                ).strip()
                if g_v2r or g_default:
                    mapping_enabled = True
                    value_to_recipients = g_v2r
                    default_recipient = g_default
                    log(
                        f"⚙ 未启用配置映射，已应用「全量默认映射」"
                        f"{len(g_v2r)} 条作为最后兜底"
                    )

        custom_enabled = bool(send.get("custom_message_enabled", False))
        custom_msg = str(send.get("custom_message", "") or "") if custom_enabled else ""
        if custom_msg:
            custom_msg = _render_weather(custom_msg, log)

        # 5) 逐人构造任务
        tasks = []
        for person in persons:
            table_data = processor.get_person_table_data(
                person, name_column, extract_columns, conditions
            )
            if not table_data:
                log(f"未找到 {person} 的数据，跳过")
                continue
            person_data = processor.format_table_data(table_data)

            if mapping_enabled and person in value_to_recipients:
                # 启用映射且命中 → 按映射展开（一对多时产生多个 task）
                recipients_list = value_to_recipients[person]
            elif mapping_enabled and default_recipient:
                # 启用映射未命中，但配置了兜底接收人 → 用兜底
                recipients_list = [default_recipient]
            else:
                # 三种场景共用此分支：
                #   1) 未启用映射（默认情况，向后兼容）
                #   2) 启用映射但未命中且无兜底 → 回退到原"筛选列"逻辑
                # 原逻辑：wechat_column 列映射 → manual_recipient → person 本身
                r = mapping.get(person) or manual_recipient or person
                recipients_list = [r]

            for recipient in recipients_list:
                tasks.append({
                    "name": person,            # 日志/确认对话框显示用筛选值
                    "person_data": person_data,
                    "table_data": table_data,
                    "recipient": recipient,    # 实际微信接收人
                    "custom_msg": custom_msg,
                })

        if not tasks:
            raise ProfilePrepareError("没有可发送的人员数据")

        attachment = str(send.get("attachment", "") or "").strip()
        if attachment and not os.path.exists(attachment):
            log(f"附加文件不存在，将不带附件发送: {attachment}")

        return {
            "title": title,
            "sheet": sheet_name,
            "tasks": tasks,
            "send_order": _resolve_send_order(send),
            "attachment": attachment,
            "send_interval": max(0.0, min(30.0, float(send.get("interval", 0.5)))),
            "chat_delay": max(0.0, min(10.0, float(send.get("chat_delay", 0.3)))),
        }
    finally:
        # 云端临时文件读完即删
        if temp_path:
            cleanup_temp_file(temp_path)
