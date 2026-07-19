from playwright.sync_api import sync_playwright
import time
import json
import logging
import os

logger = logging.getLogger(__name__)

class KDocsWebSocketCapture:
    def __init__(self):
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self.captured_messages = []
        self.debug_mode = True
    
    def _log(self, msg):
        if self.debug_mode:
            print(f"[KDocsWebSocket] {msg}")
        logger.info(msg)
    
    def _setup_browser(self):
        self._log("Setting up browser with WebSocket capture...")
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(
            headless=False,
            args=["--start-maximized"],
            timeout=60000
        )
        
        storage_file = os.path.join(os.path.dirname(__file__), "..", "browser_storage.json")
        storage_state = None
        if os.path.exists(storage_file):
            try:
                with open(storage_file, "r", encoding="utf-8") as f:
                    storage_state = json.load(f)
                self._log("Loaded saved browser storage")
            except:
                self._log("Failed to load saved browser storage")
        
        self.context = self.browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            storage_state=storage_state if storage_state else None
        )
        
        self.page = self.context.new_page()
        self.page.on("websocket", self._handle_websocket)
        self.page.on("response", self._handle_response)
    
    def _handle_websocket(self, websocket):
        self._log(f"WebSocket connected: {websocket.url}")
        
        websocket.on("framesent", lambda frame: self._log_frame("Send", frame))
        websocket.on("framereceived", lambda frame: self._log_frame("Receive", frame))
    
    def _log_frame(self, direction, frame):
        try:
            text = frame.text
            if text and len(text) > 50:
                try:
                    json_data = json.loads(text)
                    if isinstance(json_data, dict):
                        keys = list(json_data.keys())
                        if any(k in keys for k in ['cells', 'rows', 'data', 'value', 'content']):
                            self._log(f"  WS {direction}: Found table-related data")
                            self._log(f"    Keys: {keys}")
                            self.captured_messages.append({
                                "direction": direction,
                                "data": json_data
                            })
                except:
                    if len(text) > 1000:
                        self._log(f"  WS {direction}: Large message ({len(text)} chars)")
        except:
            pass
    
    def _handle_response(self, response):
        try:
            url = response.url
            if 'cell' in url.lower() or 'sheet' in url.lower():
                try:
                    content_type = response.headers.get("content-type", "")
                    if "json" in content_type:
                        json_data = response.json()
                        if isinstance(json_data, dict):
                            if 'cells' in json_data or 'data' in json_data:
                                self._log(f"  HTTP: Found data in {url}")
                                self._log(f"    Keys: {list(json_data.keys())[:10]}")
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
    
    def capture_data(self, document_url, wait_time=60):
        self._log(f"Starting WebSocket capture from: {document_url}")
        
        try:
            self._setup_browser()
            
            self._log(f"Opening URL: {document_url}")
            self.page.goto(document_url, wait_until="networkidle", timeout=120000)
            
            if not self._wait_for_login():
                self._log("Login failed")
                return None
            
            self._log(f"Waiting {wait_time} seconds for data to load...")
            time.sleep(wait_time)
            
            self._log(f"Captured {len(self.captured_messages)} WebSocket messages")
            
            if self.captured_messages:
                for i, msg in enumerate(self.captured_messages):
                    self._log(f"\n--- Message {i+1} ---")
                    self._log(f"Direction: {msg['direction']}")
                    data = msg['data']
                    if isinstance(data, dict):
                        self._log(f"Keys: {list(data.keys())[:15]}")
                        if 'cells' in data:
                            cells = data['cells']
                            if isinstance(cells, list):
                                self._log(f"Cell count: {len(cells)}")
                                if cells:
                                    self._log(f"First cell: {json.dumps(cells[0], ensure_ascii=False)[:200]}")
            
            return self.captured_messages
            
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
    
    def extract_table_from_messages(self, messages):
        for msg in messages:
            data = msg['data']
            if isinstance(data, dict):
                if 'cells' in data:
                    return self._parse_cells(data['cells'])
                
                if 'data' in data:
                    nested = data['data']
                    if isinstance(nested, dict):
                        if 'cells' in nested:
                            return self._parse_cells(nested['cells'])
        
        self._log("No table data found in WebSocket messages")
        return None
    
    def _parse_cells(self, cells):
        if not isinstance(cells, list):
            return None
        
        max_row = 0
        max_col = 0
        
        for cell in cells:
            if isinstance(cell, dict):
                r_from = cell.get('row_from', cell.get('r', 0))
                r_to = cell.get('row_to', r_from)
                c_from = cell.get('col_from', cell.get('c', 0))
                c_to = cell.get('col_to', c_from)
                
                max_row = max(max_row, r_to)
                max_col = max(max_col, c_to)
        
        table_data = [["" for _ in range(max_col + 1)] for _ in range(max_row + 1)]
        
        for cell in cells:
            if isinstance(cell, dict):
                r_from = cell.get('row_from', cell.get('r', 0))
                r_to = cell.get('row_to', r_from)
                c_from = cell.get('col_from', cell.get('c', 0))
                c_to = cell.get('col_to', c_from)
                text = cell.get('cell_text', cell.get('value', cell.get('v', '')))
                
                for r in range(r_from, r_to + 1):
                    for c in range(c_from, c_to + 1):
                        if 0 <= r <= max_row and 0 <= c <= max_col:
                            table_data[r][c] = str(text)
        
        table_data = [row for row in table_data if any(row)]
        self._log(f"Parsed {len(table_data)} rows, {len(table_data[0])} columns")
        return [table_data]

def test_websocket():
    capture = KDocsWebSocketCapture()
    url = "https://www.kdocs.cn/l/cmNZuNWSYyTy"
    
    print("=" * 60)
    print("Testing KDocs WebSocket Capture")
    print("=" * 60)
    print(f"\nURL: {url}")
    
    messages = capture.capture_data(url, wait_time=45)
    
    if messages:
        print(f"\n✓ Captured {len(messages)} WebSocket messages")
        result = capture.extract_table_from_messages(messages)
        
        if result:
            print(f"\n✓ Extracted {len(result)} tables")
            for i, table in enumerate(result):
                print(f"\nTable {i+1}: {len(table)} rows, {len(table[0])} columns")
                print("Headers:", table[0])
                print("First 3 rows:")
                for row in table[1:4]:
                    print(f"  {row}")
        else:
            print("\n✗ No table data found")
    else:
        print("\n✗ No WebSocket messages captured")

if __name__ == "__main__":
    test_websocket()