"""WPS金山文档开放平台官方API模块

文档地址: https://developer.kdocs.cn

核心流程:
1. OAuth 2.0 授权 - 获取 access_token
2. 获取文档的所有 sheet 列表
3. 获取指定 sheet 的单元格数据

使用前必须:
- 在 https://developer.kdocs.cn 注册开发者
- 创建应用并获取 AppID 和 AppKey
- 应用回调地址必须填写 http://localhost:8765/callback (与GUI默认值一致)
- 申请权限: user_basic, access_personal_files
"""

import os
import json
import time
import logging
import threading
import urllib.parse
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler

import requests

logger = logging.getLogger(__name__)

# 金山文档开放平台 API 主机
API_HOST = "https://developer.kdocs.cn"

# 令牌保存文件（使用绝对路径，确保在任何目录下运行都能找到）
TOKEN_FILE = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "wps_token.json"))

# 默认回调端口
DEFAULT_CALLBACK_PORT = 8765


class _CallbackHandler(BaseHTTPRequestHandler):
    """处理OAuth回调的HTTP请求处理器"""

    def do_GET(self):
        # 解析URL
        parsed = urllib.parse.urlparse(self.path)
        query = urllib.parse.parse_qs(parsed.query)

        if "code" in query:
            self.server.auth_code = query["code"][0]
            self.server.auth_state = query.get("state", [None])[0]
            self._send_html_response(
                "<html><body style='font-family: Microsoft YaHei; text-align: center; padding: 60px;'>"
                "<h2 style='color: green;'>✓ 授权成功！</h2>"
                "<p>已获取授权码，请返回程序窗口继续操作。</p>"
                "<p>您可以关闭此页面。</p></body></html>"
            )
            # 收到code后停止服务器
            threading.Thread(target=self.server.shutdown, daemon=True).start()
        else:
            error = query.get("error", ["unknown"])[0]
            self.server.auth_code = None
            self.server.auth_error = error
            self._send_html_response(
                f"<html><body style='font-family: Microsoft YaHei; text-align: center; padding: 60px;'>"
                f"<h2 style='color: red;'>✗ 授权失败</h2>"
                f"<p>错误: {error}</p>"
                f"<p>请返回程序窗口重试。</p></body></html>"
            )
            threading.Thread(target=self.server.shutdown, daemon=True).start()

    def _send_html_response(self, html):
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(html.encode("utf-8"))

    def log_message(self, format, *args):
        # 不打印HTTP访问日志
        pass


class WPSOfficialAPI:
    """WPS金山文档开放平台API客户端"""

    def __init__(self, app_id, app_key, redirect_uri=None):
        self.app_id = app_id
        self.app_key = app_key
        self.redirect_uri = redirect_uri or f"http://localhost:{DEFAULT_CALLBACK_PORT}/callback"
        self.token_data = self._load_token()

    # ========== 令牌管理 ==========

    def _load_token(self):
        """从本地文件加载已保存的令牌"""
        logger.info(f"Loading token from: {TOKEN_FILE}")
        if os.path.exists(TOKEN_FILE):
            try:
                with open(TOKEN_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                logger.info(f"Token file found, app_id in file: {data.get('app_id')}")
                if data.get("app_id") == self.app_id:
                    logger.info("Token loaded successfully")
                    return data
                else:
                    logger.warning(f"Token app_id mismatch: file has {data.get('app_id')}, current is {self.app_id}")
            except Exception as e:
                logger.warning(f"Failed to load token file: {e}")
        else:
            logger.info("Token file does not exist")
        return None

    def _save_token(self, token_data):
        """保存令牌到本地文件"""
        try:
            token_data["app_id"] = self.app_id
            token_data["saved_at"] = int(time.time())
            with open(TOKEN_FILE, "w", encoding="utf-8") as f:
                json.dump(token_data, f, ensure_ascii=False, indent=2)
            logger.info(f"Token saved to {TOKEN_FILE}")
        except Exception as e:
            logger.warning(f"Failed to save token: {e}")

    def _is_token_valid(self):
        """检查令牌是否有效"""
        if not self.token_data:
            return False
        access_token = self.token_data.get("access_token")
        if not access_token:
            return False
        # 检查是否过期（提前60秒）
        saved_at = self.token_data.get("saved_at", 0)
        expires_in = self.token_data.get("expires_in", 0)
        if saved_at + expires_in - 60 > time.time():
            return True
        return False

    def _get_access_token(self):
        """获取有效的 access_token，必要时自动刷新"""
        if self._is_token_valid():
            return self.token_data["access_token"]

        # 尝试刷新
        refresh_token = self.token_data.get("refresh_token") if self.token_data else None
        if refresh_token:
            logger.info("Token expired, refreshing...")
            new_token = self._refresh_token(refresh_token)
            if new_token:
                return new_token["access_token"]

        # 需要重新授权
        logger.warning("No valid token, authorization required")
        return None

    # ========== OAuth 授权 ==========

    def authorize(self, timeout=300):
        """启动OAuth授权流程

        1. 启动本地HTTP服务器接收回调
        2. 用默认浏览器打开授权页
        3. 等待用户授权
        4. 用code换取access_token
        """
        # 从回调地址解析端口
        parsed = urllib.parse.urlparse(self.redirect_uri)
        port = parsed.port or DEFAULT_CALLBACK_PORT
        path = parsed.path or "/callback"

        logger.info(f"Starting OAuth callback server on port {port}, path: {path}")

        # 启动本地HTTP服务器
        server = HTTPServer(("127.0.0.1", port), _CallbackHandler)
        server.auth_code = None
        server.auth_error = None
        server.timeout = 1

        server_thread = threading.Thread(target=server.serve_forever, daemon=True)
        server_thread.start()

        # 构造授权URL
        scope = "user_basic,access_personal_files"
        state = f"wxauto_{int(time.time())}"
        auth_url = (
            f"{API_HOST}/h5/auth"
            f"?app_id={urllib.parse.quote(self.app_id)}"
            f"&scope={urllib.parse.quote(scope)}"
            f"&redirect_uri={urllib.parse.quote(self.redirect_uri)}"
            f"&state={state}"
        )

        logger.info(f"Opening authorization URL in default browser...")
        logger.info(f"Full authorization URL: {auth_url}")
        print("\n" + "=" * 60)
        print("🔐 即将打开浏览器进行授权")
        print(f"   如果浏览器没有自动打开，请手动访问：")
        print(f"   {auth_url}")
        print("=" * 60 + "\n")

        try:
            webbrowser.open(auth_url)
        except Exception as e:
            logger.warning(f"Failed to open browser automatically: {e}")

        # 等待回调
        start_time = time.time()
        while time.time() - start_time < timeout:
            if server.auth_code is not None:
                break
            if server.auth_error is not None:
                logger.error(f"Authorization error: {server.auth_error}")
                server.shutdown()
                return None
            time.sleep(1)
            elapsed = int(time.time() - start_time)
            if elapsed % 30 == 0 and elapsed > 0:
                logger.info(f"Waiting for authorization... ({elapsed}s / {timeout}s)")
        else:
            logger.error("Authorization timeout")
            server.shutdown()
            return None

        code = server.auth_code
        logger.info(f"Got authorization code: {code[:20]}...")
        server.shutdown()

        # 用code换取 access_token
        return self._exchange_code_for_token(code)

    def _exchange_code_for_token(self, code):
        """用授权码换取 access_token"""
        url = f"{API_HOST}/api/v1/oauth2/access_token"
        params = {
            "code": code,
            "app_id": self.app_id,
            "app_key": self.app_key,
        }

        logger.info("Exchanging code for access_token...")
        logger.info(f"Request URL: {url}")
        logger.info(f"Request params (app_id masked): {{'code': '***', 'app_id': '{self.app_id}', 'app_key': '***'}}")
        try:
            resp = requests.get(url, params=params, headers={"Content-Type": "application/json"}, timeout=30)
            logger.info(f"Response status code: {resp.status_code}")
            logger.info(f"Response headers: {dict(resp.headers)}")
            
            try:
                data = resp.json()
                logger.info(f"Full token response: {data}")
            except ValueError:
                logger.info(f"Response is not JSON: {resp.text[:500]}")
                raise Exception(f"API返回非JSON格式: {resp.text[:200]}")

            if data.get("code") == 0 and data.get("data"):
                token_data = data["data"]
                self.token_data = token_data
                self._save_token(token_data)
                logger.info("Access token obtained and saved")
                return token_data
            else:
                error_code = data.get("code", "unknown")
                error_msg = data.get("msg", data.get("message", "Unknown error"))
                logger.error(f"Failed to get access_token: code={error_code}, msg={error_msg}")
                
                if error_code == 40001:
                    error_msg = f"授权失败: {error_msg}。请检查应用是否已上线，回调地址是否正确配置。"
                elif error_code == 20001:
                    error_msg = f"授权失败: {error_msg}。授权码已过期或无效，请重新授权。"
                elif error_code == 40002:
                    error_msg = f"授权失败: {error_msg}。AppID或AppKey错误。"
                
                raise Exception(f"授权失败 [错误码 {error_code}]: {error_msg}")
        except Exception as e:
            logger.error(f"Error exchanging code for token: {e}")
            raise

    def _refresh_token(self, refresh_token):
        """使用 refresh_token 刷新 access_token"""
        url = f"{API_HOST}/api/v1/oauth2/refresh_token"
        params = {"app_id": self.app_id}
        body = {
            "app_key": self.app_key,
            "refresh_token": refresh_token,
        }

        logger.info("Refreshing access_token...")
        try:
            resp = requests.post(
                url, params=params, json=body,
                headers={"Content-Type": "application/json"}, timeout=30
            )
            data = resp.json()

            if data.get("code") == 0 and data.get("data"):
                token_data = data["data"]
                self.token_data = token_data
                self._save_token(token_data)
                logger.info("Access token refreshed")
                return token_data
            else:
                logger.warning(f"Failed to refresh token: {data}")
                return None
        except Exception as e:
            logger.warning(f"Error refreshing token: {e}")
            return None

    # ========== 表格数据读取 ==========

    def _parse_file_token(self, document_url):
        """从文档URL中提取 file_token

        支持的URL格式:
        - https://www.kdocs.cn/l/{file_token}
        - https://kdocs.cn/l/{file_token}
        - https://www.kdocs.cn/l/{file_token}?...
        """
        # 提取 /l/ 后面的部分
        import re
        match = re.search(r'/l/([A-Za-z0-9]+)', document_url)
        if match:
            return match.group(1)

        # 尝试从其他常见路径提取
        match = re.search(r'/d/([A-Za-z0-9]+)', document_url)
        if match:
            return match.group(1)

        # 直接返回URL的最后一段（去掉查询参数）
        parsed = urllib.parse.urlparse(document_url)
        path_parts = parsed.path.strip('/').split('/')
        if path_parts and path_parts[-1]:
            return path_parts[-1]

        logger.warning(f"Could not parse file_token from URL: {document_url}")
        return None

    def get_sheets(self, file_token):
        """获取文档的所有 sheet 列表

        返回: list of {sheet_name, sheet_id, sheet_idx, row_from, row_to, col_from, col_to}
        """
        access_token = self._get_access_token()
        if not access_token:
            logger.error("No valid access_token, authorization required")
            return None

        url = f"{API_HOST}/api/v1/openapi/ksheet/{file_token}/sheets"
        params = {"access_token": access_token}

        logger.info(f"Getting sheets for file_token: {file_token}")
        try:
            resp = requests.get(url, params=params, timeout=30)
            data = resp.json()

            if data.get("code") == 0 and data.get("data"):
                sheets = data["data"].get("sheets_info", [])
                logger.info(f"Got {len(sheets)} sheets:")
                for s in sheets:
                    logger.info(f"  - sheet_name={s.get('sheet_name')}, sheet_id={s.get('sheet_id')}, sheet_idx={s.get('sheet_idx')}")
                    logger.info(f"    range: rows {s.get('row_from')}-{s.get('row_to')}, cols {s.get('col_from')}-{s.get('col_to')}")
                return sheets
            else:
                logger.error(f"Failed to get sheets: {data}")
                return None
        except Exception as e:
            logger.error(f"Error getting sheets: {e}")
            return None

    def get_cells(self, file_token, sheet_idx, row_from, row_to, col_from, col_to):
        """获取指定 sheet 的单元格数据

        返回: list of [row_from, row_to, col_from, col_to, cell_text, ...]
        """
        access_token = self._get_access_token()
        if not access_token:
            logger.error("No valid access_token, authorization required")
            return None

        url = f"{API_HOST}/api/v1/openapi/et/{file_token}/sheets/{sheet_idx}/cells"
        params = {
            "access_token": access_token,
            "row_from": row_from,
            "row_to": row_to,
            "col_from": col_from,
            "col_to": col_to,
        }

        logger.info(f"Getting cells: sheet_idx={sheet_idx}, rows={row_from}-{row_to}, cols={col_from}-{col_to}")
        try:
            resp = requests.get(url, params=params, timeout=60)
            data = resp.json()

            if data.get("code") == 0 and data.get("data"):
                cells = data["data"].get("cells", [])
                logger.info(f"Got {len(cells)} cells")
                return cells
            else:
                logger.error(f"Failed to get cells: {data}")
                return None
        except Exception as e:
            logger.error(f"Error getting cells: {e}")
            return None

    def extract_table(self, document_url, sheet_name=None):
        """高层接口：从WPS文档URL提取表格数据

        返回: [table_data] 格式，table_data[0] 是表头行
        """
        logger.info(f"Extracting table from: {document_url}")

        # 1. 提取 file_token
        file_token = self._parse_file_token(document_url)
        if not file_token:
            logger.error("Failed to parse file_token from URL")
            return None
        logger.info(f"File token: {file_token}")

        # 2. 检查授权
        access_token = self._get_access_token()
        if not access_token:
            logger.error("Not authorized. Please authorize first.")
            return None

        # 3. 获取 sheet 列表
        sheets = self.get_sheets(file_token)
        if not sheets:
            logger.error("Failed to get sheets or no sheets found")
            return None

        # 4. 选择目标 sheet
        target_sheet = None
        if sheet_name:
            for s in sheets:
                if s.get("sheet_name") == sheet_name:
                    target_sheet = s
                    break
            if not target_sheet:
                # 模糊匹配
                for s in sheets:
                    if sheet_name in s.get("sheet_name", ""):
                        target_sheet = s
                        logger.info(f"Fuzzy matched sheet: {s.get('sheet_name')}")
                        break
            if not target_sheet:
                logger.error(f"Sheet '{sheet_name}' not found. Available sheets: {[s.get('sheet_name') for s in sheets]}")
                return None
        else:
            target_sheet = sheets[0]
            logger.info(f"Using first sheet: {target_sheet.get('sheet_name')}")

        logger.info(f"Target sheet: {target_sheet.get('sheet_name')}")

        # 5. 获取单元格数据
        sheet_idx = target_sheet.get("sheet_idx", target_sheet.get("sheet_id", 0))
        row_from = target_sheet.get("row_from", 0)
        row_to = target_sheet.get("row_to", 1000)
        col_from = target_sheet.get("col_from", 0)
        col_to = target_sheet.get("col_to", 50)

        # 安全上限：避免一次拉取过多数据
        max_rows = 5000
        max_cols = 200
        if row_to - row_from > max_rows:
            row_to = row_from + max_rows
            logger.warning(f"Limiting rows to {max_rows}")
        if col_to - col_from > max_cols:
            col_to = col_from + max_cols
            logger.warning(f"Limiting cols to {max_cols}")

        cells = self.get_cells(file_token, sheet_idx, row_from, row_to, col_from, col_to)
        if not cells:
            logger.error("Failed to get cells or empty result")
            return None

        # 6. 把单元格数据组装成二维表
        table_data = self._cells_to_table(cells, row_from, row_to, col_from, col_to)
        if table_data:
            logger.info(f"Extracted table: {len(table_data)} rows, {len(table_data[0]) if table_data else 0} cols")
            return [table_data]
        return None

    def _cells_to_table(self, cells, row_from, row_to, col_from, col_to):
        """把API返回的单元格数据组装成二维表

        每个cell包含: row_from, row_to, col_from, col_to, cell_text, origin_cell_value, num_format
        """
        if not cells:
            return None

        # 计算表格大小
        actual_row_to = row_from
        actual_col_to = col_from
        for cell in cells:
            if cell.get("row_to", 0) > actual_row_to:
                actual_row_to = cell["row_to"]
            if cell.get("col_to", 0) > actual_col_to:
                actual_col_to = cell["col_to"]

        rows = actual_row_to - row_from + 1
        cols = actual_col_to - col_from + 1

        if rows <= 0 or cols <= 0:
            logger.warning(f"Invalid table size: rows={rows}, cols={cols}")
            return None

        # 初始化空表
        table = [["" for _ in range(cols)] for _ in range(rows)]

        # 填充单元格
        filled = 0
        for cell in cells:
            cr_from = cell.get("row_from", 0)
            cc_from = cell.get("col_from", 0)
            cr_to = cell.get("row_to", cr_from)
            cc_to = cell.get("col_to", cc_from)

            text = cell.get("cell_text", "")
            if text is None:
                text = ""
            text = str(text)

            # 处理合并单元格：填充到选区内的所有单元格
            for r in range(cr_from - row_from, cr_to - row_from + 1):
                for c in range(cc_from - col_from, cc_to - col_from + 1):
                    if 0 <= r < rows and 0 <= c < cols:
                        if not table[r][c]:  # 不覆盖已有内容
                            table[r][c] = text
            filled += 1

        logger.info(f"Filled {filled} cells into {rows}x{cols} table")

        # 移除全空的尾部行和列
        table = self._trim_empty_rows_cols(table)
        return table

    def _trim_empty_rows_cols(self, table):
        """移除表尾的全空行和全空列"""
        if not table:
            return table

        # 找最后一行有数据的
        last_data_row = -1
        for i, row in enumerate(table):
            if any(cell.strip() if isinstance(cell, str) else cell for cell in row):
                last_data_row = i

        if last_data_row < 0:
            return table

        table = table[:last_data_row + 1]

        # 找最后一列有数据的
        max_col = 0
        for row in table:
            for j in range(len(row) - 1, -1, -1):
                cell = row[j]
                if cell and (cell.strip() if isinstance(cell, str) else cell):
                    if j + 1 > max_col:
                        max_col = j + 1
                    break

        table = [row[:max_col] for row in table]
        return table


def test_api():
    """简单测试"""
    print("WPS Official API Test")
    print("=" * 60)

    app_id = input("AppID: ").strip()
    app_key = input("AppKey: ").strip()
    redirect_uri = "http://localhost:8765/callback"

    api = WPSOfficialAPI(app_id, app_key, redirect_uri)

    # 授权
    if not api._is_token_valid():
        print("\n需要授权...")
        token = api.authorize()
        if not token:
            print("授权失败")
            return

    # 提取表格
    url = input("\n文档URL: ").strip()
    sheet_name = input("Sheet名称(可选): ").strip() or None

    result = api.extract_table(url, sheet_name)
    if result:
        table = result[0]
        print(f"\n✓ 提取成功: {len(table)} 行, {len(table[0])} 列")
        print("表头:", table[0])
        print("\n前3行数据:")
        for row in table[1:4]:
            print(" ", row)
    else:
        print("\n✗ 提取失败")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
    test_api()
