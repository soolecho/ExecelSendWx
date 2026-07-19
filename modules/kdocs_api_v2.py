from playwright.sync_api import sync_playwright
import requests
import json
import logging
import time
import re
import os

logger = logging.getLogger(__name__)

class KDocsAPIV2:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Origin": "https://www.kdocs.cn",
            "Referer": "https://www.kdocs.cn/"
        })
        self.session_token = None
        self.file_token = None
    
    def _extract_file_token(self, document_url):
        match = re.search(r'/l/([a-zA-Z0-9]+)', document_url)
        if match:
            return match.group(1)
        match = re.search(r'/file/([a-zA-Z0-9]+)', document_url)
        if match:
            return match.group(1)
        return None
    
    def _get_session_token_via_browser(self, document_url):
        self._log("Getting session token via browser...")
        
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=False, timeout=60000)
                
                storage_file = os.path.join(os.path.dirname(__file__), "..", "browser_storage.json")
                storage_state = None
                if os.path.exists(storage_file):
                    try:
                        with open(storage_file, "r", encoding="utf-8") as f:
                            storage_state = json.load(f)
                        self._log("Loaded saved browser storage")
                    except:
                        pass
                
                context = browser.new_context(
                    viewport={"width": 1920, "height": 1080},
                    storage_state=storage_state if storage_state else None
                )
                
                page = context.new_page()
                page.on("response", self._handle_response_for_token)
                
                self._log(f"Opening URL: {document_url}")
                page.goto(document_url, wait_until="networkidle", timeout=120000)
                
                timeout = 120
                start_time = time.time()
                while time.time() - start_time < timeout:
                    if self.session_token:
                        self._log(f"Got session token!")
                        break
                    
                    try:
                        body_text = page.inner_text("body")
                        if len(body_text) > 2000:
                            self._log("Login detected, waiting for token...")
                    except:
                        pass
                    
                    time.sleep(3)
                
                try:
                    storage_state = context.storage_state()
                    with open(storage_file, "w", encoding="utf-8") as f:
                        json.dump(storage_state, f, ensure_ascii=False, indent=2)
                    self._log("Saved browser storage")
                except:
                    pass
                
                browser.close()
                
                return self.session_token
                
        except Exception as e:
            self._log(f"Error getting session token: {e}")
            return None
    
    def _handle_response_for_token(self, response):
        try:
            url = response.url
            if '/api/v3/office/session/' in url and 'et?first' in url:
                try:
                    json_data = response.json()
                    if isinstance(json_data, dict) and 'token' in json_data:
                        self.session_token = json_data['token']
                        self._log(f"Found session token: {self.session_token[:20]}...")
                except:
                    pass
        except:
            pass
    
    def _log(self, msg):
        print(f"[KDocsAPI] {msg}")
        logger.info(msg)
    
    def _get_api_data(self, path, params=None):
        if not self.session_token:
            self._log("No session token, cannot call API")
            return None
        
        base_url = "https://www.kdocs.cn"
        url = f"{base_url}{path}"
        
        if params is None:
            params = {}
        
        params["token"] = self.session_token
        
        try:
            response = self.session.get(url, params=params)
            self._log(f"API call: {url} -> {response.status_code}")
            
            if response.status_code == 200:
                try:
                    result = response.json()
                    return result
                except:
                    self._log(f"Response is not JSON: {response.text[:200]}")
            else:
                self._log(f"API error: {response.status_code}")
                
        except Exception as e:
            self._log(f"API call failed: {e}")
        
        return None
    
    def get_sheets(self):
        if not self.file_token:
            self._log("No file token")
            return None
        
        path = f"/api/v3/office/file/{self.file_token}/sheets"
        result = self._get_api_data(path)
        
        if result and result.get("code") == 0 and "data" in result:
            sheets = result["data"]
            self._log(f"Found {len(sheets)} sheets")
            for sheet in sheets:
                self._log(f"  - Sheet {sheet['index']}: {sheet['name']}")
            return sheets
        
        return None
    
    def get_sheet_data(self, sheet_idx=0, row_from=0, row_to=1000, col_from=0, col_to=50):
        if not self.file_token:
            self._log("No file token")
            return None
        
        path = f"/api/v3/office/file/{self.file_token}/sheets/{sheet_idx}/cells"
        params = {
            "row_from": row_from,
            "row_to": row_to,
            "col_from": col_from,
            "col_to": col_to
        }
        
        result = self._get_api_data(path, params)
        
        if result and result.get("code") == 0 and "data" in result:
            cells = result["data"].get("cells", [])
            self._log(f"Found {len(cells)} cell ranges")
            
            max_row = row_to
            max_col = col_to
            table_data = [["" for _ in range(max_col - col_from + 1)] for _ in range(max_row - row_from + 1)]
            
            for cell in cells:
                r_from = cell["row_from"] - row_from
                r_to = cell["row_to"] - row_from
                c_from = cell["col_from"] - col_from
                c_to = cell["col_to"] - col_from
                
                if 0 <= r_from <= r_to < len(table_data) and 0 <= c_from <= c_to < len(table_data[0]):
                    for r in range(r_from, min(r_to + 1, len(table_data))):
                        for c in range(c_from, min(c_to + 1, len(table_data[0]))):
                            table_data[r][c] = cell.get("cell_text", "")
            
            table_data = [row for row in table_data if any(row)]
            self._log(f"Extracted {len(table_data)} rows, {len(table_data[0])} columns")
            return [table_data]
        
        return None
    
    def extract_full_sheet(self, sheet_name=None, max_rows=1000, max_cols=50):
        sheets = self.get_sheets()
        
        if not sheets:
            self._log("No sheets found, trying default")
            return self.get_sheet_data(0, 0, max_rows, 0, max_cols)
        
        sheet_idx = 0
        if sheet_name:
            for sheet in sheets:
                if sheet.get("name") == sheet_name:
                    sheet_idx = sheet.get("index", 0)
                    break
            else:
                self._log(f"Sheet '{sheet_name}' not found, using first")
        
        self._log(f"Extracting sheet {sheet_idx}")
        return self.get_sheet_data(sheet_idx, 0, max_rows, 0, max_cols)
    
    def extract_from_url(self, document_url, sheet_name=None):
        self.file_token = self._extract_file_token(document_url)
        if not self.file_token:
            self._log("Cannot extract file token")
            return None
        
        self._log(f"File token: {self.file_token}")
        
        token = self._get_session_token_via_browser(document_url)
        if not token:
            self._log("Failed to get session token")
            return None
        
        return self.extract_full_sheet(sheet_name)

def test_api_v2():
    api = KDocsAPIV2()
    url = "https://www.kdocs.cn/l/cmNZuNWSYyTy"
    
    print("=" * 60)
    print("Testing KDocs API V2")
    print("=" * 60)
    print(f"\nURL: {url}")
    
    result = api.extract_from_url(url, "7.15投诉")
    
    if result:
        print(f"\n✓ Extraction successful!")
        for i, table in enumerate(result):
            print(f"\nTable {i+1}: {len(table)} rows, {len(table[0])} columns")
            print("Headers:", table[0])
            print("First 3 rows:")
            for row in table[1:4]:
                print(f"  {row}")
    else:
        print("\n✗ Extraction failed")

if __name__ == "__main__":
    test_api_v2()