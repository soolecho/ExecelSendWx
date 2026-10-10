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


# ---------------------------------------------------------------- 指定区域截图（关闭对比模式）
def parse_region(spec: str) -> Optional[Tuple[int, int, int, int]]:
    """解析 Excel 区域 "A1:F20"（大小写不敏感）→ (r0, r1, c0, c1)。

    返回 0-based 半开区间 [r0, r1) x [c0, c1)，可直接切片二维网格；
    终点单元格为包含式（"A1:F20" = 第 1~20 行、A~F 列）。
    只写单格 "F20" 视为 1x1；spec 为空/非法返回 None（调用方按整表已用区域处理）。
    """
    s = str(spec or "").strip().upper()
    if not s:
        return None
    m = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", s)
    if not m:
        return None

    def col_idx(letters: str) -> int:
        n = 0
        for ch in letters:
            n = n * 26 + (ord(ch) - 64)
        return n - 1

    c0, r0 = col_idx(m.group(1)), int(m.group(2)) - 1
    if m.group(3):
        # Excel 范围终点是包含式的："A1:F20" 含第 1~20 行、A~F 列
        c1, r1 = col_idx(m.group(3)) + 1, int(m.group(4))
    else:
        c0, c1 = col_idx(m.group(1)), col_idx(m.group(1)) + 1
        r1 = r0 + 1
    if r0 < 0 or c0 < 0 or r1 <= r0 or c1 <= c0:
        return None
    return r0, r1, c0, c1


def slice_region(grid: List[List[str]], r0: int, r1: int, c0: int, c1: int) -> List[List[str]]:
    """按 0-based 半开区间从整表网格切出矩形区域，缺列补空串保证等宽。"""
    width = c1 - c0
    out: List[List[str]] = []
    for row in grid[r0:r1]:
        seg = [str(v) if v is not None else "" for v in row[c0:c1]]
        seg += [""] * (width - len(seg))
        out.append(seg)
    return out


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


def _disp_width(text) -> int:
    """显示宽度：中文/全角按 2，半角按 1（用于列宽估算）。"""
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(text))


def render_rows_image(
    headers: List[str],
    rows: List[List[str]],
    out_png: str,
    log_fn=None,
) -> Optional[str]:
    """用 QPainter 直接绘制表格为 PNG（不依赖 QtWebEngine，worker 线程可用）。

    特性：
    - 列宽按表头/内容自适应（自动展开，无需手动拉宽）
    - 超长单元格自动折行、行高自适应，保证内容发全
    返回生成路径；异常返回 None。
    """
    try:
        from PyQt6.QtCore import QRect, Qt
        from PyQt6.QtGui import (QColor, QFont, QFontMetrics, QImage, QPainter,
                                 QPen)

        data = [list(r) for r in (rows or [])]
        hds = [str(h) for h in (headers or [])]
        ncol = max(len(hds), max((len(r) for r in data), default=0))
        if ncol == 0:
            return None
        hds = (hds + [""] * ncol)[:ncol]
        data = [(r + [""] * ncol)[:ncol] for r in data]

        font = QFont("Microsoft YaHei", 10)
        fm = QFontMetrics(font)
        pad = 8
        max_col_px = 520          # 超长列折行宽度上限
        min_col_px = 60

        # 列宽：表头与内容实际像素宽度（含 padding），超长封顶后自动折行
        col_w = []
        for c in range(ncol):
            w = max(fm.horizontalAdvance(hds[c]) + pad * 2, min_col_px)
            for r in data:
                w = max(w, fm.horizontalAdvance(str(r[c])) + pad * 2)
            col_w.append(min(w, max_col_px))

        def _wrap(text: str, width: int) -> list:
            t = str(text)
            if not t:
                return [""]
            lines, cur = [], ""
            for ch in t:
                if cur and fm.horizontalAdvance(cur + ch) > width - pad * 2:
                    lines.append(cur)
                    cur = ch
                else:
                    cur += ch
            lines.append(cur)
            return lines

        cells = [[_wrap(str(r[c]), col_w[c]) for c in range(ncol)] for r in data]
        header_h = max(fm.height() + pad * 2, 34)
        row_h = [
            max(fm.lineSpacing() * len(cells[i][c]) + pad * 2 for c in range(ncol))
            for i in range(len(data))
        ]
        total_w = sum(col_w) + 1
        total_h = header_h + sum(row_h) + 1

        img = QImage(total_w, total_h, QImage.Format.Format_ARGB32)
        img.fill(QColor("white"))
        p = QPainter(img)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        p.setFont(font)
        grid_pen = QPen(QColor("#d0d0d0"))
        header_bg = QColor("#2F6FEB")
        alt_bg = QColor("#F7FAFF")

        # 表头
        p.fillRect(0, 0, total_w, header_h, header_bg)
        p.setPen(QPen(QColor("white")))
        x = 0
        for c, w in enumerate(col_w):
            p.drawText(QRect(x + pad, 0, w - pad * 2, header_h),
                       Qt.AlignmentFlag.AlignVCenter, hds[c])
            x += w

        # 数据行（自动折行绘制）
        y = header_h
        for i, r in enumerate(cells):
            if i % 2 == 1:
                p.fillRect(0, y, total_w, row_h[i], alt_bg)
            p.setPen(QPen(QColor("#333333")))
            x = 0
            for c, w in enumerate(col_w):
                ty = y
                for ln in r[c]:
                    p.drawText(QRect(x + pad, ty, w - pad * 2, fm.lineSpacing()),
                               Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                               ln)
                    ty += fm.lineSpacing()
                x += w
            y += row_h[i]

        # 网格线
        p.setPen(grid_pen)
        x = 0
        for c, w in enumerate(col_w):
            x += w
            p.drawLine(x, 0, x, total_h)
        y = header_h
        for h in row_h:
            y += h
            p.drawLine(0, y, total_w, y)
        p.end()

        ok = img.save(out_png)
        return out_png if ok and os.path.exists(out_png) else None
    except Exception as exc:
        (log_fn or (lambda m: None))(f"图片渲染失败，跳过图片推送：{exc}")
        return None


# ---------------------------------------------------------------- 数据清洗/校验（阶段B）
def apply_clean_rules(
    headers: List[str],
    rows: List[List[str]],
    rules,
) -> Tuple[List[List[str]], List[dict]]:
    """按每任务的清洗/校验规则加工行。

    每列加工顺序：去空格 -> 去空 -> 去非数字 -> 取前N位 -> 前缀校验。
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
            if rule.drop_empty and not val.strip():
                # 去空(空剔除)：该列清洗后为空 → 整行剔除，不受 on_fail 影响。
                # on_fail 只决定"校验失败"（前缀等）时剔/留；空值属数据质量过滤，
                # 若受 on_fail="keep" 门控，空行会混进推送/智能表格（实测空账号行被粘贴）。
                dropped.append(
                    {"row_idx": r_idx, "column": rule.column,
                     "reason": f"内容为空(清洗后): {raw}"}
                )
                keep_row = False
                break
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
                # on_fail="keep"：行保留，但该列仍应应用清洗变换（val 已去非数字/取前N位），
                # 否则前缀失败的行会以原始未清洗值写入推送/智能表格（如 12 位原始账号而非前 11 位）
                if idx < len(row_out):
                    row_out[idx] = val
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
    # 自动调整列宽：按表头与内容最大显示宽度（中文按2宽），上限 50、下限 8
    for c in range(len(sub_headers)):
        w = _disp_width(sub_headers[c])
        for r in sub_rows:
            w = max(w, _disp_width(r[c]))
        ws.column_dimensions[openpyxl.utils.get_column_letter(c + 1)].width = \
            min(max(w + 2, 8), 50)
    try:
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
    except OSError:
        pass
    wb.save(out_path)
    return out_path


# ---------------------------------------------------------------- AirScript 智能表格写入
_AIRSCRIPT_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)


def _col_letter(n: int) -> str:
    """1 → A, 26 → Z, 27 → AA ..."""
    n = int(n or 1)
    if n < 1:
        n = 1
    s = ""
    while n > 0:
        rem = (n - 1) % 26
        s = chr(65 + rem) + s
        n = (n - 1) // 26
    return s


def airsync_append(
    webhook: str,
    token: str,
    sheet: str,
    start_col: int,
    col_count: int,
    rows: List[List[str]],
    timeout: int = 30,
) -> Tuple[bool, str]:
    """把清洗/提取后的多行数据追加写入智能表格（AirScript 通用工具库脚本）。

    脚本需为「文档共享脚本」并内置 wps_airscript_client_api.js（Http 工具库，
    支持 function= getUsedRangeData / setRangeValues）。流程分两步：
      1) getUsedRangeData 取已用区域 [startRow, startCol, nRows, nCols] → 计算末尾空行
      2) setRangeValues 把 rows（已去掉首行标题）从末尾空行写入指定列起
    rows 为空时跳过写入；返回 (成功?, 详情)。
    """
    import requests

    if not webhook or not token:
        return False, "未配置 AirScript webhook 或脚本令牌"
    if not rows:
        return True, "无可写入行（0 行），跳过"

    # 按列数剪裁、统一转字符串
    data_rows: List[List[str]] = []
    for r in rows:
        src = r[:col_count] if col_count > 0 else r
        data_rows.append([str(c) for c in src])
    if not data_rows:
        return True, "剪裁后无可写入行，跳过"

    headers = {
        "Content-Type": "application/json",
        "AirScript-Token": token,
        "User-Agent": _AIRSCRIPT_UA,
    }

    def _call(argv: dict) -> tuple:
        """POST 一次 webhook，返回 (ok, result_payload)。"""
        try:
            resp = requests.post(
                webhook,
                json={"Context": {"argv": argv}},
                headers=headers,
                timeout=timeout,
            )
        except Exception as exc:
            return False, {"http_error": repr(exc)}
        if resp.status_code != 200:
            return False, {"http_status": resp.status_code, "body": resp.text[:200]}
        try:
            body = resp.json()
        except ValueError:
            return True, {"non_json": resp.text[:120]}
        if body.get("error"):
            return False, {"script_error": body["error"]}
        return True, body

    # ---------- 第 1 步：读已用区域，定位末尾 ----------
    ok, body = _call({"function": "getUsedRangeData", "isGetData": False})
    if not ok:
        return False, f"读取表格末尾失败: {body}"
    if body.get("non_json"):
        return False, f"读取表格末尾失败: {body['non_json']}"
    result = body.get("data", {}).get("result") or []
    used = None
    if isinstance(result, list) and result:
        item = result[0]
        d = item.get("data") if isinstance(item, dict) else None
        if isinstance(d, list) and len(d) >= 4:
            used = d  # [startRow, startCol, nRows, nCols]
    if not used:
        return False, f"读取表格末尾失败: 无法解析 getUsedRangeData 响应 {result}"

    start_row, start_col_used, used_rows, used_cols = (int(x) for x in used[:4])
    # 末尾空行 = 起始行 + 已用行数（即最后一个非空行的下一行）
    insert_row = start_row + used_rows
    col_begin = int(start_col or 1)
    if col_begin < start_col_used:
        # 用户指定的起始列可能在工作表已用区域的左侧，直接以指定列为准
        col_begin = max(int(start_col or 1), 1)
    n_cols = len(data_rows[0])
    col_end = col_begin + n_cols - 1

    # ---------- 第 2 步：批量写入（写值 + 跟随上方行格式） ----------
    address = f"{_col_letter(col_begin)}{insert_row}:{_col_letter(col_end)}{insert_row + len(data_rows) - 1}"
    ok, body = _call(
        {
            "function": "setRangeValuesWithFormat",
            "address": address,
            "values": data_rows,
            "thisSheetName": sheet or "",
            "formatFromRow": max(start_row, insert_row - 1),  # 参考格式行 = 最后一非空行
        }
    )
    if not ok:
        return False, f"写入失败: {body}"
    if body.get("non_json"):
        return False, f"写入失败: {body['non_json']}"
    result2 = body.get("data", {}).get("result") or []
    ok_flag = bool(
        isinstance(result2, list)
        and result2
        and (isinstance(result2[0], dict) and result2[0].get("success"))
        and not (isinstance(result2[0], dict) and result2[0].get("error"))
    )
    if not ok_flag:
        # 脚本端未实现 setRangeValuesWithFormat（旧脚本库）→ 回退纯写值
        ok_fb, body_fb = _call(
            {
                "function": "setRangeValues",
                "address": address,
                "values": data_rows,
                "thisSheetName": sheet or "",
            }
        )
        if not ok_fb:
            return False, f"写入失败: {body_fb}"
        if body_fb.get("non_json"):
            return False, f"写入失败: {body_fb['non_json']}"
        result3 = body_fb.get("data", {}).get("result") or []
        ok_flag = bool(
            isinstance(result3, list)
            and result3
            and (isinstance(result3[0], dict) and result3[0].get("success"))
        )
        if not ok_flag:
            return False, f"脚本未确认写入成功: {result3}"
        return True, f"追加 {len(data_rows)} 行到 {address}（旧脚本，无格式跟随）"
    return True, f"追加 {len(data_rows)} 行到 {address}（含格式跟随）"