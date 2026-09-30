"""监控处理纯逻辑（无 QWidget、可单测）。

职责：
1. 读取指定 sheet → 筛选出要的数据行
2. 增量对比：与上次已处理行键集合对比，只取"新增行"
3. 生成推送内容：文字说明 / 图片 / 生成的表格文件
4. 关闭对比时：按文件指纹判定"新文件"
"""
import hashlib
import os
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
) -> str:
    """生成“新增了哪些”的文字说明（可读表格排版）。"""
    if not new_rows:
        return ""
    lines = [f"共新增 {len(new_rows)} 条："]
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


def append_to_ledger(
    headers: List[str],
    rows: List[List[str]],
    ledger_path: str,
    sheet_name: str = "",
    remark_col: str = "",
    remark_text: str = "",
    exclude_extra: bool = True,
    clean_rules=None,
    log_fn=None,
) -> int:
    """把清洗后的行追加到本地台账 xlsx，返回写入行数。

    目标 sheet：sheet_name 为空取第一个；文件不存在/无该 sheet 则新建。
    写入方式：跳过已有数据行，从目标列首个空行起逐行下移（整行追加到最底部）。
    exclude_extra=True 时只用 clean_rules 涉及的列；否则按 header 同名匹配写入
    （台账缺列时自动补列）。remark_col 非空时在同样新增行的该列写备注
    （remark_text 为空则用"已追加 <时间>"）。原子写（tmp + os.replace）。
    """
    import os
    import tempfile
    import openpyxl

    log = log_fn or (lambda msg: None)
    if not rows:
        return 0
    try:
        os.makedirs(os.path.dirname(os.path.abspath(ledger_path)), exist_ok=True)
    except OSError:
        pass

    wb = None
    ws = None
    if os.path.exists(ledger_path):
        try:
            wb = openpyxl.load_workbook(ledger_path)
        except Exception as exc:
            log(f"台账文件加载失败({exc})，将新建覆盖")
            wb = None
    if wb is None:
        wb = openpyxl.Workbook()
        ws = wb.active
        if sheet_name:
            ws.title = sheet_name
    else:
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        elif not sheet_name:
            ws = wb.active
        else:
            ws = wb.create_sheet(sheet_name)
    if ws is None:
        ws = wb.active

    # 用 clean_rules 涉及的列名集合；exclude_extra=False 时按 header 同名写
    from .monitor_config import CleanRule

    rule_cols = set()
    for rule in clean_rules or []:
        if isinstance(rule, CleanRule) and rule.column:
            rule_cols.add(rule.column)
    wanted_set = rule_cols if (exclude_extra and rule_cols) else set(
        c for c in headers if c
    )
    # 按源表表头顺序排序目标列，保证新建台账列顺序与源一致
    wanted_cols = [c for c in headers if c in wanted_set]

    # 解析台账表头（行1）与列位置
    ledger_headers = [
        str(ws.cell(row=1, column=col).value or "").strip() for col in range(1, ws.max_column + 1)
    ]
    has_header = any(ledger_headers)
    write_col: Dict[str, int] = {}
    if not has_header:
        # 空台账：目标列从第 1 列起按源顺序写表头
        for i, col_name in enumerate(wanted_cols, start=1):
            ws.cell(row=1, column=i, value=col_name)
            write_col[col_name] = i
        next_col = len(wanted_cols) + 1
    else:
        next_col = ws.max_column + 1
        for col_name in wanted_cols:
            try:
                write_col[col_name] = ledger_headers.index(col_name) + 1
            except ValueError:
                # 台账缺列 → 追加为最后一列
                write_col[col_name] = next_col
                ws.cell(row=1, column=write_col[col_name], value=col_name)
                next_col += 1
    if remark_col:
        if not has_header or remark_col not in ledger_headers:
            remark_col_idx = next_col
            ws.cell(row=1, column=remark_col_idx, value=remark_col)
        else:
            remark_col_idx = ledger_headers.index(remark_col) + 1
    else:
        remark_col_idx = None

    # 定位最底部（第一个全空行）→ 从该行起逐行写入
    start_row = 1
    while True:
        empty = all(
            (ws.cell(row=start_row, column=col).value in (None, ""))
            for col in range(1, ws.max_column + 1)
        )
        if empty:
            break
        start_row += 1
        if start_row > 100000:
            break

    remark_value = (remark_text or "").strip() or f"已追加 {_now_str()}"
    for i, row in enumerate(rows):
        target = start_row + i
        for col_name, col_idx in write_col.items():
            src_idx = headers.index(col_name) if col_name in headers else -1
            if src_idx >= 0 and src_idx < len(row):
                ws.cell(row=target, column=col_idx, value=row[src_idx])
        if remark_col_idx:
            ws.cell(row=target, column=remark_col_idx, value=remark_value)

    # 原子写
    fd, tmp = tempfile.mkstemp(
        prefix=".ledger.", suffix=".xlsx", dir=os.path.dirname(os.path.abspath(ledger_path))
    )
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, ledger_path)
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise
    return len(rows)


def _now_str() -> str:
    from datetime import datetime

    return datetime.now().strftime("%Y-%m-%d %H:%M")


def build_out_file(
    headers: List[str],
    new_rows: List[List[str]],
    out_path: str,
) -> str:
    """把新增行写入一个临时 xlsx，返回文件路径。"""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for r in new_rows:
        ws.append(r)
    try:
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
    except OSError:
        pass
    wb.save(out_path)
    return out_path