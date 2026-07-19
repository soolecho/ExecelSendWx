import requests
import json
import logging
import time
import re

logger = logging.getLogger(__name__)

class KDocsAPI:
    def __init__(self):
        self.base_url = "https://www.kdocs.cn"
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"
        })
    
    def _extract_file_token(self, document_url):
        match = re.search(r'/l/([a-zA-Z0-9]+)', document_url)
        if match:
            return match.group(1)
        
        match = re.search(r'/file/([a-zA-Z0-9]+)', document_url)
        if match:
            return match.group(1)
        
        logger.warning(f"Cannot extract file token from URL: {document_url}")
        return None
    
    def _get_session_token(self, file_token):
        try:
            url = f"{self.base_url}/api/v3/office/session/{file_token}/et?first"
            response = self.session.get(url)
            result = response.json()
            
            if "token" in result:
                logger.info(f"Got session token: {result['token'][:20]}...")
                return result
            
            logger.error(f"Failed to get session token: {result}")
            return None
            
        except Exception as e:
            logger.error(f"Error getting session token: {e}")
            return None
    
    def _get_collaborator_data(self, file_token):
        try:
            url = f"{self.base_url}/kfc/miniprovider/v1/links/collaborator"
            params = {"fid": file_token}
            response = self.session.get(url, params=params)
            result = response.json()
            
            if result.get("code") == 0 and "data" in result:
                logger.info(f"Got collaborator data with {len(result['data'])} items")
                return result["data"]
            
            logger.error(f"Failed to get collaborator data: {result}")
            return None
            
        except Exception as e:
            logger.error(f"Error getting collaborator data: {e}")
            return None
    
    def _get_et_api(self, file_token, path):
        try:
            session_info = self._get_session_token(file_token)
            if not session_info:
                return None
            
            token = session_info["token"]
            region = session_info.get("region", "cn")
            replica = session_info.get("replica", 0)
            
            url = f"https://{region}-{replica}-api.wps.cn/office/api/v3/{path}"
            params = {"token": token}
            
            response = self.session.get(url, params=params)
            
            if response.status_code == 200:
                result = response.json()
                logger.info(f"ET API {path} success")
                return result
            
            logger.error(f"ET API {path} failed: {response.status_code}")
            return None
            
        except Exception as e:
            logger.error(f"Error calling ET API {path}: {e}")
            return None
    
    def get_sheets(self, file_token):
        try:
            url = f"{self.base_url}/api/v3/office/file/{file_token}/sheets"
            response = self.session.get(url)
            result = response.json()
            
            if result.get("code") == 0 and "data" in result:
                sheets = result["data"]
                logger.info(f"Found {len(sheets)} sheets")
                for sheet in sheets:
                    logger.info(f"  - Sheet {sheet['index']}: {sheet['name']}")
                return sheets
            
            logger.error(f"Failed to get sheets: {result}")
            
            data = self._get_collaborator_data(file_token)
            if data and isinstance(data, dict) and "sheets" in data:
                sheets = data["sheets"]
                logger.info(f"Found {len(sheets)} sheets from collaborator API")
                return sheets
            
            return None
            
        except Exception as e:
            logger.error(f"Error getting sheets: {e}")
            return None
    
    def get_sheet_data(self, file_token, sheet_idx=0, row_from=0, row_to=1000, col_from=0, col_to=50):
        try:
            url = f"{self.base_url}/api/v3/office/file/{file_token}/sheets/{sheet_idx}/cells"
            params = {
                "row_from": row_from,
                "row_to": row_to,
                "col_from": col_from,
                "col_to": col_to
            }
            
            response = self.session.get(url, params=params)
            result = response.json()
            
            if result.get("code") == 0 and "data" in result:
                cells = result["data"].get("cells", [])
                logger.info(f"Found {len(cells)} cell ranges")
                
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
                logger.info(f"Extracted {len(table_data)} rows, {len(table_data[0])} columns")
                return [table_data]
            
            logger.error(f"Failed to get sheet data: {result}")
            return None
            
        except Exception as e:
            logger.error(f"Error getting sheet data: {e}")
            return None
    
    def extract_full_sheet(self, file_token, sheet_name=None, max_rows=1000, max_cols=50):
        sheets = self.get_sheets(file_token)
        if not sheets:
            logger.warning("No sheets found, trying default sheet")
            return self.get_sheet_data(file_token, 0, 0, max_rows, 0, max_cols)
        
        sheet_idx = 0
        if sheet_name:
            for sheet in sheets:
                if sheet.get("name") == sheet_name or sheet.get("title") == sheet_name:
                    sheet_idx = sheet.get("index", 0)
                    break
            else:
                logger.warning(f"Sheet '{sheet_name}' not found, using first sheet")
        
        logger.info(f"Extracting sheet {sheet_idx}")
        
        return self.get_sheet_data(file_token, sheet_idx, 0, max_rows, 0, max_cols)
    
    def extract_from_url(self, document_url, sheet_name=None):
        file_token = self._extract_file_token(document_url)
        if not file_token:
            logger.error("Cannot extract file token from URL")
            return None
        
        logger.info(f"File token: {file_token}")
        
        return self.extract_full_sheet(file_token, sheet_name)

def test_api():
    api = KDocsAPI()
    url = "https://www.kdocs.cn/l/cmNZuNWSYyTy"
    print(f"\nTesting API extraction for: {url}")
    
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
    logging.basicConfig(level=logging.INFO)
    test_api()