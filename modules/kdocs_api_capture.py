from playwright.sync_api import sync_playwright
import time
import json
import logging
import os

logger = logging.getLogger(__name__)

class KDocsAPICapture:
    def __init__(self):
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self.captured_data = []
        self.debug_mode = True
    
    def _log(self, msg):
        if self.debug_mode:
            print(f"[KDocsAPICapture] {msg}")
        logger.info(msg)
    
    def _setup_browser(self):
        self._log("Setting up browser with network interception...")
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(
            headless=False,
            args=["--start-maximized"],
            timeout=60000
        )
        
        self.context = self.browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        
        storage_file = os.path.join(os.path.dirname(__file__), "..", "browser_storage.json")
        if os.path.exists(storage_file):
            try:
                with open(storage_file, "r", encoding="utf-8") as f:
                    storage_state = json.load(f)
                    self.context.storage_state(storage_state=storage_state)
                self._log("Loaded saved browser storage")
            except:
                self._log("Failed to load saved browser storage")
        
        self.page = self.context.new_page()
        self.page.on("response", self._handle_response)
    
    def _handle_response(self, response):
        try:
            url = response.url
            
            interesting_keywords = [
                'cell', 'sheet', 'data', 'grid', 'rows', 'columns',
                'document', 'workbook', 'kdocs', 'wps', 'collaborator',
                'et/', 'entry', 'content', 'table', 'range'
            ]
            
            if any(keyword in url.lower() for keyword in interesting_keywords):
                try:
                    content_type = response.headers.get("content-type", "")
                    if "json" in content_type or "text/plain" in content_type:
                        try:
                            json_data = response.json()
                            if isinstance(json_data, dict) or isinstance(json_data, list):
                                data_str = json.dumps(json_data, ensure_ascii=False)
                                if len(data_str) > 100:
                                    self._log(f"✓ Captured potential data from: {url}")
                                    self._log(f"  Size: {len(data_str)} bytes")
                                    
                                    if isinstance(json_data, dict):
                                        self._log(f"  Keys: {list(json_data.keys())[:10]}")
                                    else:
                                        self._log(f"  List length: {len(json_data)}")
                                    
                                    self.captured_data.append({
                                        "url": url,
                                        "data": json_data,
                                        "size": len(data_str)
                                    })
                        except:
                            try:
                                text = response.text
                                if len(text) > 100:
                                    self._log(f"✓ Captured text from: {url}")
                                    self._log(f"  Size: {len(text)} bytes")
                            except:
                                pass
                except:
                    pass
        except:
            pass
    
    def _wait_for_login(self, timeout=120):
        self._log(f"Waiting for login, timeout: {timeout} seconds")
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            try:
                body_text = self.page.inner_text("body")
                body_len = len(body_text)
                if body_len > 2000:
                    self._log(f"Login detected! Body text length: {body_len}")
                    return True
            except:
                pass
            
            time.sleep(3)
            elapsed = int(time.time() - start_time)
            if elapsed % 30 == 0:
                self._log(f"Waiting... ({elapsed}s / {timeout}s)")
        
        self._log("Login timeout")
        return False
    
    def capture_api_data(self, document_url, wait_time=60):
        self._log(f"Starting API capture from: {document_url}")
        
        try:
            self._setup_browser()
            
            self._log(f"Opening URL: {document_url}")
            self.page.goto(document_url, wait_until="networkidle", timeout=120000)
            
            if not self._wait_for_login():
                self._log("Login failed")
                return None
            
            self._log(f"Waiting {wait_time} seconds for data to load...")
            time.sleep(wait_time)
            
            self._log(f"Captured {len(self.captured_data)} API responses")
            
            if self.captured_data:
                for i, item in enumerate(self.captured_data):
                    self._log(f"\n--- Response {i+1} ---")
                    self._log(f"URL: {item['url']}")
                    self._log(f"Size: {item['size']} bytes")
                    self._log(f"Data keys: {list(item['data'].keys())[:15]}")
                    
                    if 'cells' in item['data']:
                        self._log(f"✓ Found cells data!")
                        cells = item['data']['cells']
                        if isinstance(cells, list):
                            self._log(f"  Cell count: {len(cells)}")
                            if cells:
                                self._log(f"  First cell: {json.dumps(cells[0], ensure_ascii=False)[:200]}")
                    
                    if 'rows' in item['data']:
                        self._log(f"✓ Found rows data!")
                        rows = item['data']['rows']
                        if isinstance(rows, list):
                            self._log(f"  Row count: {len(rows)}")
            
            return self.captured_data
            
        except Exception as e:
            self._log(f"Capture failed: {e}")
            logger.error(f"Capture failed: {e}", exc_info=True)
            return None
        finally:
            if self.playwright:
                try:
                    self.playwright.stop()
                except:
                    pass
    
    def extract_table_from_api(self, captured_data):
        for item in captured_data:
            data = item['data']
            
            if isinstance(data, dict):
                if 'cells' in data:
                    return self._parse_cells_data(data['cells'])
                
                if 'rows' in data:
                    return self._parse_rows_data(data['rows'])
                
                if 'data' in data:
                    nested_data = data['data']
                    if isinstance(nested_data, dict):
                        if 'cells' in nested_data:
                            return self._parse_cells_data(nested_data['cells'])
                        if 'rows' in nested_data:
                            return self._parse_rows_data(nested_data['rows'])
        
        self._log("No table data found in captured responses")
        return None
    
    def _parse_cells_data(self, cells):
        if not isinstance(cells, list):
            return None
        
        max_row = 0
        max_col = 0
        
        for cell in cells:
            if isinstance(cell, dict):
                row_from = cell.get('row_from', 0)
                row_to = cell.get('row_to', 0)
                col_from = cell.get('col_from', 0)
                col_to = cell.get('col_to', 0)
                
                max_row = max(max_row, row_to)
                max_col = max(max_col, col_to)
        
        table_data = [["" for _ in range(max_col + 1)] for _ in range(max_row + 1)]
        
        for cell in cells:
            if isinstance(cell, dict):
                r_from = cell.get('row_from', 0)
                r_to = cell.get('row_to', 0)
                c_from = cell.get('col_from', 0)
                c_to = cell.get('col_to', 0)
                text = cell.get('cell_text', '')
                
                for r in range(r_from, r_to + 1):
                    for c in range(c_from, c_to + 1):
                        if 0 <= r <= max_row and 0 <= c <= max_col:
                            table_data[r][c] = text
        
        table_data = [row for row in table_data if any(row)]
        self._log(f"Parsed {len(table_data)} rows, {len(table_data[0])} columns")
        return [table_data]
    
    def _parse_rows_data(self, rows):
        if not isinstance(rows, list):
            return None
        
        table_data = []
        for row in rows:
            if isinstance(row, list):
                table_data.append([str(cell) if cell is not None else "" for cell in row])
            elif isinstance(row, dict):
                row_data = []
                for key in sorted(row.keys()):
                    row_data.append(str(row[key]) if row[key] is not None else "")
                table_data.append(row_data)
        
        if table_data:
            self._log(f"Parsed {len(table_data)} rows")
            return [table_data]
        
        return None

def test_capture():
    capture = KDocsAPICapture()
    url = "https://www.kdocs.cn/l/cmNZuNWSYyTy"
    
    print("=" * 60)
    print("Testing KDocs API Capture")
    print("=" * 60)
    print(f"\nURL: {url}")
    
    captured_data = capture.capture_api_data(url, wait_time=30)
    
    if captured_data:
        print(f"\n✓ Captured {len(captured_data)} API responses")
        result = capture.extract_table_from_api(captured_data)
        
        if result:
            print(f"\n✓ Extracted {len(result)} tables")
            for i, table in enumerate(result):
                print(f"\nTable {i+1}: {len(table)} rows, {len(table[0])} columns")
                print("Headers:", table[0])
                print("First 3 rows:")
                for row in table[1:4]:
                    print(f"  {row}")
        else:
            print("\n✗ No table data found in API responses")
    else:
        print("\n✗ No API responses captured")

if __name__ == "__main__":
    test_capture()