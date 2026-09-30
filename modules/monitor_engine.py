"""监控处理纯逻辑（无 QWidget、可单测）。

职责：
1. 读取指定 sheet → 筛选出要的数据行
2. 增量对比：与上次已处理行键集合对比，只取"新增行"
3. 生成推送内容：文字说明 / 图片 / 生成的表格文件
4. 关闭对比时：按文件指纹判定"新文件"
"""
import hashlib
import os
import re
from typing import Dict, List, Optional, Tuple

from . import wps_cloud


# ---------------------------------------------------------------- 读取与筛选
def read_sheet_rows(
    file_path: str,
    sheet_name: str = "",
    log_fn=None,
) -> Tuple[List[str], List[List[str]]]:
    """读取指定 sheet → (表头, 数据行二维数组)。
    sheet_name 为空时取第一个 sheet；失败抛异常由调用方处理。"""
    log = log_fn or (lambda msg: None)
    names = wps_cloud.read_sheet_names(file_path, log_fn=log)
    if not names:
        raise ValueError("无法读取工作簿的 sheet 名称（文件可能损坏或不是 Excel）")
    target = sheet_name if sheet_name and sheet_name in names else names[0]
    rows = wps_cloud.read_one_sheet(file_path, target, log_fn=log)
    if rows is None:
        raise ValueError(f"Sheet「{target}」读取失败")
    if not rows:
        return [], []
    headers = [str(c) for c in rows[0]]
    data = rows[1:]
    return headers, data


def filter_rows(
    headers: List[str],
    rows: List[List[str]],
    filter_column: str = "",
    filter_values: List[str] | None = None,
) -> List[List[str]]:
    """按单列命中过滤行；filter_column 为空返回全部行。"""
    if not filter_column:
        return [list(r) for r in rows]
    filters = set(filter_values or [])
    if not filters:
        return [list(r) for r in rows]
    try:
        idx = headers.index(filter_column)
    except ValueError:
        # 列不存在：精确优先，兜底子串匹配表头（与项目标题列匹配一致）
        idx = next(
            (i for i, h in enumerate(headers) if filter_column in str(h)),
            -1,
        )
        if idx < 0:
            return []
    result = []
    for r in rows:
        cell = str(r[idx]) if idx < len(r) else ""
        if cell in filters:
            result.append(list(r))
    return result


# ---------------------------------------------------------------- 增量对比
def row_key(msg: object) -> str:
    """任意对象 → 稳定字符串指纹（用于行键）。"""
    return hashlib.sha256(str(msg).encode("utf-8", "ignore")).hexdigest()


def diff_incremental(
    headers: List[str],
    rows: List[List[str]],
    baseline_keys: List[str],
    key_columns: List[str] | None = None,
) -> Tuple[List[List[str]], List[str]]:
    """返回 (新增行, 新增行键)。行键取 key_columns 对应列值；为空则整行哈希。
    已存在的行键不再重复推送。"""
    baseline = set(baseline_keys or [])
    new_rows: List[List[str]] = []
    new_keys: List[str] = []
    indexes: List[int] = []
    if key_columns:
        for c in key_columns:
            try:
                indexes.append(headers.index(c))
            except ValueError:
                indexes.append(next((i for i, h in enumerate(headers) if c in str(h)), -1))
        indexes = [i for i in indexes if i >= 0]

    for r in rows:
        if indexes:
            parts = [str(r[i]) if i < len(r) else "" for i in indexes]
            key = row_key("|".join(parts))
        else:
            key = row_key("|".join(str(c) for c in r))
        if key in baseline:
            continue
        new_keys.append(key)
        new_rows.append(list(r))
    return new_rows, new_keys


def build_baseline(
    headers: List[str],
    rows: List[List[str]],
    key_columns: List[str] | None = None,
) -> List[str]:
    """把当前全部行的键写回基线（用于刷新 baseline）。"""
    _, keys = diff_incremental(headers, rows, [], key_columns)
    return keys


# ---------------------------------------------------------------- 文件指纹（关闭对比）
def file_fingerprint(file_path: str) -> str:
    """文件的轻量指纹（size+mtime_ns），用于判定目ut新文件。"""
    try:
        st = os.stat(file_path)
        return f"{st.st_size}:{st.st_mtime_ns}"
    except OSError:
        return ""


def is_new_file(seen: Dict[str, str], file_path: str) -> bool:
    """关闭对比模式：文件名未记录或指纹变化 → 视为新文件。"""
    name = os.path.basename(file_path)
    fp = file_fingerprint(file_path)
    return seen.get(name) != fp


# ---------------------------------------------------------------- 内容生成
def render_rows_text(
    headers: List[str],
    new_rows: List[List[str]],
    title: str = "",
    show_detail: bool = True,
) -> str:
    """生成推送文字说明。

    title 为空用默认标题；show_detail=True 时在下方附逐行明细（列:值），
    False 时仅显示标题 + 新增条数。
    """
    if not new_rows:
        return ""
    t = str(title).strip() or "监控新增提醒"
    lines = [f"{t}", f"共新增 {len(new_rows)} 条"]
    if show_detail:
        lines.append("")
        # 用表头前 4 列 + 后续值截断，避免过长
        cols = headers[:4]
        for r in new_rows:
            parts = [f"{cols[i]}:{r[i]}" for i in range(len(cols)) if i < len(r)]
            extra = len(r) - len(cols)
            if extra > 0:
                parts.append(f"…+{extra}列")
            lines.append("  " + "，".join(parts))
    return "\n".join(lines)


def render_html(headers: List[str], rows: List[List[str]]) -> str:
    """生成带样式的 HTML 表格（供图片渲染/文件导出复用）。"""
    cells = ["<tr><th>序号</th>"] + [f"<th>{_esc(h)}</th>" for h in headers] + ["</tr>"]
    body = []
    for i, r in enumerate(rows, start=1):
        body.append("<tr><td>{}</td>".format(i) + "".join(
            f"<td>{_esc(str(c))}</td>" for c in r
        ) + "</tr>")
    html = (
        "<html><head><meta charset='utf-8'><style>"
        "table{border-collapse:collapse;font-family:'Microsoft YaHei',sans-serif;font-size:13px}"
        "th,td{border:1px solid #ccc;padding:5px 9px;text-align:left}"
        "th{background:#f0f0f0}</style></head><body><table>"
        + "".join(cells) + "".join(body) + "</table></body></html>"
    )
    return html


def _esc(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def render_rows_image(
    headers: List[str],
    rows: List[List[str]],
    out_png: str,
    log_fn=None,
) -> Optional[str]:
    """尽力把表格渲染成 PNG。返回成功时的路径；环境不支持则返回 None（调用方自动跳过图片推送）。"""
    log = log_fn or (lambda msg: None)
    html = render_html(headers, rows)
    try:
        from PyQt6.QtCore import QEventLoop, QTimer, QUrl
        from PyQt6.QtWebEngineWidgets import QWebEngineView  # type: ignore
        from PyQt6.QtWidgets import QApplication

        QApplication.instance() or QApplication([])
        view = QWebEngineView()
        view.resize(1000, 200 + len(rows) * 30)
        view.load(QUrl.fromLocalFile(_write_tmp_html(html, out_png)))

        loop = QEventLoop()
        image_holder = {}

        def on_ready():
            QTimer.singleShot(150, _grab)

        def _grab():
            image_holder["img"] = view.grab()
            loop.quit()

        view.page().loadFinished.connect(on_ready)
        timer = QTimer()
        timer.setSingleShot(True)
        timer.timeout.connect(loop.quit)
        timer.start(10000)
        loop.exec()
        img = image_holder.get("img")
        if img is not None:
            img.save(out_png)
        return out_png if os.path.exists(out_png) else None
    except Exception as exc:
        log(f"图片渲染不可用，跳过图片推送：{exc}")
        return None


def _write_tmp_html(html: str, out_png: str) -> str:
    tmp = out_png + ".html"
    with open(tmp, "w", encoding="utf-8") as fp:
        fp.write(html)
    return str(tmp).replace("\\", "/")


# ---------------------------------------------------------------- 数据清洗/校验（阶段B）
def apply_clean_rules(
    headers: List[str],
    rows: List[List[str]],
    rules,
) -> Tuple[List[List[str]], List[dict]]:
    """按每任务的清洗/校验规则加工行。

    每列加工顺序：去空 -> 去非数字 -> 取前N位 -> 前缀校验。
    on_fail="drop" 该行剔除；on_fail="keep" 保留清洗后的值但不拦该行。
    返回 (清洗后保留的行, 剔除信息列表 [{row_idx, column, reason}]).
    """
    from .monitor_config import CleanRule

    if not rules:
        return [list(r) for r in rows], []
    # 解析列位置（精确优先，兜底子串匹配，与 filter_rows 一致）
    col_index: Dict[int, CleanRule] = {}
    for rule in rules:
        if not rule or not rule.column:
            continue
        try:
            idx = headers.index(rule.column)
        except ValueError:
            idx = next(
                (i for i, h in enumerate(headers) if rule.column in str(h)), -1
            )
        if idx >= 0:
            col_index[idx] = rule

    kept: List[List[str]] = []
    dropped: List[dict] = []
    for r_idx, row in enumerate(rows):
        row_out = list(row)
        keep_row = True
        for idx, rule in col_index.items():
            raw = str(row_out[idx]) if idx < len(row_out) else ""
            val = raw
            if rule.strip_space:
                val = val.replace(" ", "").replace("\u3000", "")
            if rule.keep_digits_only:
                val = "".join(ch for ch in val if ch.isdigit())
            if rule.take_first_n > 0:
                val = val[: rule.take_first_n]
            if rule.require_prefix and not val.startswith(rule.require_prefix):
                dropped.append(
                    {"row_idx": r_idx, "column": rule.column,
                     "reason": f"前缀不匹配(需 {rule.require_prefix}): {raw}"}
                )
                if rule.on_fail == "drop":
                    keep_row = False
                    break
                # on_fail="keep" 保留原值，但记录提示
                continue
            if idx < len(row_out):
                row_out[idx] = val
        if keep_row:
            kept.append(row_out)
    return kept, dropped


def extract_columns_by(
    headers: List[str],
    rows: List[List[str]],
    extract_columns: List[str] | None = None,
):
    """按指定的提取列裁剪，返回 (sub_headers, sub_rows)。
    extract_columns 为空 -> 返回原始。列名精确优先，兜底子串匹配
    （与 filter_rows / apply_clean_rules 一致）。"""
    if not extract_columns:
        return list(headers), [list(r) for r in rows]
    indexes: List[int] = []
    for c in extract_columns:
        if not c:
            continue
        try:
            idx = headers.index(c)
        except ValueError:
            idx = next((i for i, h in enumerate(headers) if c in str(h)), -1)
        if idx >= 0 and idx not in indexes:
            indexes.append(idx)
    sub_headers = [headers[i] for i in indexes]
    sub_rows = [[row[i] for i in indexes if i < len(row)] for row in rows]
    return sub_headers, sub_rows


def _coerce_numeric(value: object) -> object:
    """纯数字字符串 → 数值类型（int/float），避免 Excel「文本型数字」绿三角。"""
    if isinstance(value, str):
        v = value.strip()
        if re.fullmatch(r"-?\d+(\.\d+)?", v):
            try:
                return int(v) if "." not in v else float(v)
            except (ValueError, OverflowError):
                return value
    return value


def build_out_file(
    headers: List[str],
    new_rows: List[List[str]],
    out_path: str,
    extract_columns: List[str] | None = None,
) -> str:
    """把新增行（可选按提取列裁剪后）写入一个临时 xlsx，返回文件路径。"""
    import openpyxl
    sub_headers, sub_rows = extract_columns_by(headers, new_rows, extract_columns)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(sub_headers)
    for r in sub_rows:
        ws.append([_coerce_numeric(c) for c in r])
    try:
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
    except OSError:
        pass
    wb.save(out_path)
    return out_path