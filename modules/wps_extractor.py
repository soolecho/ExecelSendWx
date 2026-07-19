from playwright.sync_api import sync_playwright
import re
import time
import logging
import os
import json
import traceback

logger = logging.getLogger(__name__)

STORAGE_STATE_FILE = "browser_storage.json"


class WPSExtractor:
    def __init__(self, headless=False):
        self.headless = headless
        self.browser = None
        self.page = None
        self.playwright = None
        self.context = None

    def _setup_browser(self):
        logger.info("=" * 50)
        logger.info("Setting up browser...")
        logger.info(f"Headless mode: {self.headless}")
        
        try:
            self.playwright = sync_playwright().start()
            logger.info("Playwright started successfully")
            
            self.browser = self.playwright.chromium.launch(
                headless=self.headless,
                args=["--start-maximized", "--disable-gpu", "--no-sandbox"],
                timeout=60000
            )
            logger.info("Browser launched successfully")
            
            storage_state = None
            if os.path.exists(STORAGE_STATE_FILE):
                try:
                    with open(STORAGE_STATE_FILE, "r", encoding="utf-8") as f:
                        storage_state = json.load(f)
                    logger.info(f"Loaded saved browser storage state from {STORAGE_STATE_FILE}")
                except Exception as e:
                    logger.warning(f"Failed to load storage state: {e}")
                    storage_state = None
            else:
                logger.info(f"No storage state file found at {STORAGE_STATE_FILE}")
            
            context_options = {
                "viewport": {"width": 1920, "height": 1080},
                "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            }
            
            if storage_state:
                context_options["storage_state"] = storage_state
                logger.info("Will use saved storage state for browser context")
            
            self.context = self.browser.new_context(**context_options)
            self.page = self.context.new_page()
            logger.info("Browser context and page created successfully")
            
            self.page.on("pageerror", lambda err: logger.error(f"Page error: {err}"))
            self.page.on("console", lambda msg: logger.debug(f"Console: {msg.text}"))
            
        except Exception as e:
            logger.error(f"Failed to setup browser: {e}")
            logger.error(traceback.format_exc())
            raise

    def _save_storage_state(self):
        try:
            if self.context:
                storage_state = self.context.storage_state()
                with open(STORAGE_STATE_FILE, "w", encoding="utf-8") as f:
                    json.dump(storage_state, f, ensure_ascii=False, indent=2)
                logger.info(f"Browser storage state saved to {STORAGE_STATE_FILE}")
        except Exception as e:
            logger.warning(f"Failed to save storage state: {e}")

    def open_wps_document(self, document_url):
        logger.info(f"Opening WPS document: {document_url}")
        self._setup_browser()
        
        try:
            self.page.goto(document_url, wait_until="networkidle", timeout=120000)
            logger.info("Document page loaded successfully")
        except Exception as e:
            logger.error(f"Failed to load document: {e}")
            logger.error(traceback.format_exc())
            raise

    def wait_for_login(self, timeout=300):
        logger.info(f"Waiting for login, timeout: {timeout} seconds")
        start_time = time.time()
        last_status = None
        
        while time.time() - start_time < timeout:
            try:
                is_logged_in = self._is_logged_in()
                
                if is_logged_in:
                    logger.info("✅ Login detected successfully")
                    self._save_storage_state()
                    return True
                
                if last_status != "login_page":
                    logger.info("⏳ Login page detected, please log in manually in the browser window...")
                    last_status = "login_page"
                
            except Exception as e:
                logger.debug(f"Checking login status: {e}")
            
            time.sleep(3)
            
            elapsed = time.time() - start_time
            if int(elapsed) % 30 == 0:
                remaining = int(timeout - elapsed)
                logger.info(f"⏳ Still waiting for login... ({int(elapsed)}s / {timeout}s, {remaining}s remaining)")
                try:
                    current_url = self.page.url
                    logger.info(f"Current URL: {current_url}")
                except:
                    pass
        
        logger.error(f"❌ Login timeout after {timeout} seconds")
        logger.info("Please check if:")
        logger.info("1. The browser window is visible")
        logger.info("2. You have completed the login")
        logger.info("3. The WPS document page has loaded")
        return False

    def _is_logged_in(self):
        try:
            current_url = self.page.url
            logger.info(f"Current page URL: {current_url}")
        except:
            current_url = ""
        
        document_selectors = [
            ".wps-doc-container",
            ".wps-spreadsheet-container",
            "[class*='doc-container']",
            "[class*='sheet-container']",
            ".wps-editor",
            "[class*='wps-content']",
            ".wps-office-container",
            "[class*='office-container']",
            "#wps-content",
            ".wps-canvas",
            "[role='document']",
            "[class*='grid-canvas']",
            ".wps-grid-container",
            ".wps-spreadsheet",
            "[class*='spreadsheet']",
            ".wps-table"
        ]
        
        for selector in document_selectors:
            try:
                element = self.page.query_selector(selector)
                if element:
                    logger.info(f"✓ Document detected via selector: {selector}")
                    return True
            except:
                pass
        
        try:
            content = self.page.content()
            doc_keywords = ["wps-doc-container", "wps-spreadsheet-container", "wps-office"]
            if any(keyword in content for keyword in doc_keywords):
                logger.info("✓ Document detected via page content")
                return True
        except:
            pass
        
        try:
            if "/doc/" in current_url or "/sheet/" in current_url or "/wps.cn" in current_url:
                body_text = self.page.inner_text("body")
                if len(body_text) > 2000:
                    logger.info(f"✓ Document page detected via content length: {len(body_text)} characters")
                    return True
        except:
            pass
        
        try:
            canvas_elements = self.page.query_selector_all("canvas, [class*='canvas']")
            if len(canvas_elements) > 3:
                logger.info(f"✓ Document canvas detected: {len(canvas_elements)} canvas elements")
                return True
        except:
            pass
        
        try:
            grid_elements = self.page.query_selector_all("[role='grid'], [class*='grid']")
            if len(grid_elements) > 0:
                logger.info(f"✓ Grid elements detected: {len(grid_elements)}")
                return True
        except:
            pass
        
        try:
            login_keywords = ["登录", "login", "password", "验证码", "account"]
            body_text = self.page.inner_text("body").lower()
            login_element_count = 0
            
            password_inputs = self.page.query_selector_all("input[type='password']")
            login_element_count += len(password_inputs)
            
            login_buttons = self.page.query_selector_all("button:has-text('登录'), button:has-text('Login')")
            login_element_count += len(login_buttons)
            
            if login_element_count > 0:
                logger.info(f"✗ Login page detected: {login_element_count} login-specific elements")
                return False
            
            has_login_keyword = any(keyword in body_text for keyword in login_keywords)
            if has_login_keyword and len(body_text) < 5000:
                logger.info(f"✗ Possible login page: contains login keywords and short content")
                return False
        except Exception as e:
            logger.debug(f"Login check error: {e}")
        
        try:
            self.page.wait_for_selector("body", timeout=5000)
            body_text = self.page.inner_text("body")
            if len(body_text) > 1000:
                logger.info(f"✓ Page loaded with substantial content: {len(body_text)} chars")
                return True
        except:
            pass
        
        logger.warning("⚠ Unable to determine login status, assuming logged in")
        return True

    def extract_table_data(self, retries=3):
        logger.info("Starting table data extraction")
        
        try:
            self._debug_page_structure()
        except Exception as e:
            logger.error(f"Debug analysis failed: {e}")
        
        for attempt in range(retries + 1):
            try:
                logger.info(f"Extraction attempt {attempt + 1}/{retries + 1}")
                
                time.sleep(3)
                
                result = self._extract_tables_via_playwright()
                if result:
                    return result
                
                result = self._extract_tables_via_javascript()
                if result:
                    return result
                
                result = self._extract_tables_via_html()
                if result:
                    return result
                
                logger.warning(f"No tables found on attempt {attempt + 1}")
                
            except Exception as e:
                logger.error(f"Failed to extract table data on attempt {attempt + 1}: {e}")
                logger.error(traceback.format_exc())
            
            if attempt < retries:
                logger.info(f"Retrying extraction... ({attempt + 1}/{retries})")
                time.sleep(5)
        
        logger.info("Trying WPS-specific extraction methods...")
        
        result = self._extract_from_iframes()
        if result:
            return result
        
        result = self._extract_wps_sheet_data()
        if result:
            return result
        
        result = self._extract_via_text_parsing()
        if result:
            return result
        
        logger.error("All extraction attempts failed")
        return []

    @staticmethod
    def extract_via_download(document_url, sheet_name=None):
        logger.info(f"=" * 60)
        logger.info("Trying download method for extraction")
        logger.info(f"URL: {document_url}")
        logger.info(f"Sheet: {sheet_name}")
        logger.info("=" * 60)
        
        try:
            from modules.kdocs_downloader import download_and_extract
            result = download_and_extract(document_url, sheet_name)
            
            if result:
                logger.info(f"Download extraction successful!")
                return result
            else:
                logger.error("Download extraction failed")
                return None
                
        except Exception as e:
            logger.error(f"Download extraction error: {e}", exc_info=True)
            return None

    def _extract_via_text_parsing(self):
        logger.info("Method 6: Text-based parsing")
        
        try:
            all_text = self.page.inner_text("body")
            logger.info(f"Total text length: {len(all_text)} characters")
            
            if len(all_text) < 100:
                logger.info("Not enough text on page")
                return None
            
            lines = all_text.split('\n')
            lines = [line.strip() for line in lines if line.strip()]
            
            logger.info(f"Total lines: {len(lines)}")
            logger.info(f"First 20 lines:")
            for line in lines[:20]:
                logger.info(f"  - {line}")
            
            if len(lines) < 5:
                logger.info("Not enough lines to form a table")
                return None
            
            table_data = []
            max_columns = 0
            
            for i, line in enumerate(lines):
                cells = line.split('\t')
                
                if len(cells) == 1:
                    cells = line.split('  ')
                    cells = [c.strip() for c in cells if c.strip()]
                
                if len(cells) == 1:
                    cells = line.split(' | ')
                
                if len(cells) > max_columns:
                    max_columns = len(cells)
                
                if len(cells) >= 2:
                    table_data.append(cells)
            
            if len(table_data) >= 2 and max_columns >= 2:
                logger.info(f"Text parsing successful: {len(table_data)} rows, {max_columns} columns")
                logger.info(f"First 3 rows:")
                for row in table_data[:3]:
                    logger.info(f"  - {row[:5]}...")
                return [table_data]
            
            logger.info("Text parsing did not produce valid table structure")
            return None
            
        except Exception as e:
            logger.error(f"Method 6 failed: {e}")
            return None

    def _extract_from_iframes(self):
        logger.info("Method 4: Trying to extract from iframes")
        
        try:
            js_script = """
            var allData = [];
            
            var iframes = document.querySelectorAll('iframe');
            logger.info('Found ' + iframes.length + ' iframes');
            
            for (var i = 0; i < iframes.length; i++) {
                try {
                    var iframeDoc = iframes[i].contentDocument || iframes[i].contentWindow.document;
                    var tables = iframeDoc.querySelectorAll('table');
                    logger.info('Iframe ' + i + ': ' + tables.length + ' tables');
                    
                    for (var j = 0; j < tables.length; j++) {
                        var rows = tables[j].querySelectorAll('tr');
                        var tableData = [];
                        for (var k = 0; k < rows.length; k++) {
                            var cells = rows[k].querySelectorAll('td, th');
                            var rowData = [];
                            for (var l = 0; l < cells.length; l++) {
                                rowData.push(cells[l].textContent.trim());
                            }
                            if (rowData.length > 0) {
                                tableData.push(rowData);
                            }
                        }
                        if (tableData.length > 0) {
                            allData.push(tableData);
                        }
                    }
                } catch(e) {
                    logger.info('Cannot access iframe ' + i + ': ' + e.message);
                }
            }
            
            return allData;
            """
            
            result = self.page.evaluate(js_script)
            
            if result and isinstance(result, list) and len(result) > 0:
                logger.info(f"Method 4 successful: {len(result)} tables from iframes")
                return result
            
            logger.info("Method 4 returned empty result from iframes")
            return None
            
        except Exception as e:
            logger.error(f"Method 4 failed: {e}")
            return None

    def _extract_wps_sheet_data(self):
        logger.info("Method 5: WPS-specific data extraction")
        
        try:
            js_script = """
            function extractWPSData() {
                var data = [];
                
                var gridContainer = document.querySelector('.wps-grid-container, [class*="grid-container"], [class*="sheet-content"]');
                if (gridContainer) {
                    var rows = gridContainer.querySelectorAll('[role="row"], .wps-row, [class*="row-"]');
                    for (var i = 0; i < rows.length; i++) {
                        var cells = rows[i].querySelectorAll('[role="gridcell"], [role="columnheader"], .wps-cell, [class*="cell-"]');
                        var rowData = [];
                        for (var j = 0; j < cells.length; j++) {
                            var cellText = cells[j].textContent || cells[j].innerText || '';
                            rowData.push(cellText.trim());
                        }
                        if (rowData.length > 0 && rowData.some(function(t) { return t.length > 0; })) {
                            data.push(rowData);
                        }
                    }
                }
                
                if (data.length === 0) {
                    var allDivs = document.querySelectorAll('div');
                    var potentialTables = [];
                    for (var i = 0; i < allDivs.length; i++) {
                        var div = allDivs[i];
                        var childDivs = div.querySelectorAll('div');
                        if (childDivs.length > 10) {
                            var gridRows = div.querySelectorAll('[class*="row"]');
                            if (gridRows.length > 5) {
                                var tableData = [];
                                for (var j = 0; j < gridRows.length; j++) {
                                    var gridCells = gridRows[j].querySelectorAll('[class*="cell"]');
                                    var rowData = [];
                                    for (var k = 0; k < gridCells.length; k++) {
                                        rowData.push(gridCells[k].textContent.trim());
                                    }
                                    if (rowData.length > 0) {
                                        tableData.push(rowData);
                                    }
                                }
                                if (tableData.length > 5) {
                                    potentialTables.push(tableData);
                                }
                            }
                        }
                    }
                    
                    if (potentialTables.length > 0) {
                        data = potentialTables[0];
                    }
                }
                
                return data;
            }
            return extractWPSData();
            """
            
            result = self.page.evaluate(js_script)
            
            if result and isinstance(result, list) and len(result) > 0:
                logger.info(f"Method 4 successful: {len(result)} rows")
                for i, row in enumerate(result[:5]):
                    logger.info(f"Row {i+1}: {row[:5]}...")
                return [result]
            
            logger.info("Method 4 returned empty result")
            return None
            
        except Exception as e:
            logger.error(f"Method 4 failed: {e}")
            return None

    def _extract_tables_via_playwright(self):
        logger.info("Method 1: Extracting tables via Playwright selectors")
        
        tables = self.page.query_selector_all("table")
        logger.info(f"Found {len(tables)} tables on page")
        
        if len(tables) == 0:
            logger.info("No tables found with 'table' selector, trying other selectors...")
            tables = self.page.query_selector_all("[role='grid'], [class*='grid'], [class*='table']")
            logger.info(f"Found {len(tables)} elements with alternative selectors")
        
        table_data = []
        
        for table_idx, table in enumerate(tables):
            try:
                rows = table.query_selector_all("tr, [role='row']")
                logger.info(f"Table {table_idx + 1}: {len(rows)} rows")
                
                table_rows = []
                for row in rows:
                    cells = row.query_selector_all("td, th, [role='gridcell'], [role='columnheader']")
                    row_data = [cell.inner_text().strip() for cell in cells]
                    
                    if row_data:
                        table_rows.append(row_data)
                
                if table_rows:
                    table_data.append(table_rows)
                    logger.info(f"Table {table_idx + 1}: extracted {len(table_rows)} rows")
            except Exception as e:
                logger.error(f"Error processing table {table_idx + 1}: {e}")
        
        if table_data:
            total_rows = sum(len(t) for t in table_data)
            logger.info(f"Method 1 successful: {len(table_data)} tables, {total_rows} total rows")
            return table_data
        
        return None

    def _extract_tables_via_javascript(self):
        logger.info("Method 2: Extracting tables via JavaScript")
        
        try:
            js_script = """
            function extractTables() {
                var tables = [];
                var allTables = document.querySelectorAll('table');
                for (var i = 0; i < allTables.length; i++) {
                    var table = allTables[i];
                    var rows = table.querySelectorAll('tr');
                    var tableData = [];
                    for (var j = 0; j < rows.length; j++) {
                        var cells = rows[j].querySelectorAll('td, th');
                        var rowData = [];
                        for (var k = 0; k < cells.length; k++) {
                            rowData.push(cells[k].textContent.trim());
                        }
                        if (rowData.length > 0) {
                            tableData.push(rowData);
                        }
                    }
                    if (tableData.length > 0) {
                        tables.push(tableData);
                    }
                }
                
                if (tables.length === 0) {
                    var grids = document.querySelectorAll('[role=\"grid\"], [class*=\"grid\"]');
                    for (var i = 0; i < grids.length; i++) {
                        var grid = grids[i];
                        var rows = grid.querySelectorAll('[role=\"row\"]');
                        var tableData = [];
                        for (var j = 0; j < rows.length; j++) {
                            var cells = rows[j].querySelectorAll('[role=\"gridcell\"], [role=\"columnheader\"]');
                            var rowData = [];
                            for (var k = 0; k < cells.length; k++) {
                                rowData.push(cells[k].textContent.trim());
                            }
                            if (rowData.length > 0) {
                                tableData.push(rowData);
                            }
                        }
                        if (tableData.length > 0) {
                            tables.push(tableData);
                        }
                    }
                }
                
                return tables;
            }
            return extractTables();
            """
            
            result = self.page.evaluate(js_script)
            
            if result and isinstance(result, list) and len(result) > 0:
                logger.info(f"Method 2 successful: {len(result)} tables")
                for i, table in enumerate(result):
                    logger.info(f"Table {i+1}: {len(table)} rows")
                return result
            
            logger.info("Method 2 returned empty result")
            return None
            
        except Exception as e:
            logger.error(f"Method 2 failed: {e}")
            return None

    def _extract_tables_via_html(self):
        logger.info("Method 3: Extracting tables via HTML parsing")
        
        try:
            html = self.page.content()
            logger.info(f"Page HTML length: {len(html)} characters")
            
            import re
            table_pattern = re.compile(r'<table[^>]*>(.*?)</table>', re.DOTALL)
            tables = table_pattern.findall(html)
            
            if len(tables) > 0:
                logger.info(f"Method 3 found {len(tables)} tables")
                
                result = []
                for table_html in tables:
                    rows = re.findall(r'<tr[^>]*>(.*?)</tr>', table_html, re.DOTALL)
                    table_data = []
                    for row_html in rows:
                        cells = re.findall(r'<(td|th)[^>]*>(.*?)</\1>', row_html, re.DOTALL)
                        row_data = [re.sub(r'<[^>]+>', '', cell[1]).strip() for cell in cells]
                        if row_data and any(row_data):
                            table_data.append(row_data)
                    if table_data:
                        result.append(table_data)
                        logger.info(f"Table extracted: {len(table_data)} rows")
                
                if result:
                    logger.info(f"Method 3 successful: {len(result)} tables")
                    return result
            
            logger.info("Method 3 returned empty result")
            return None
            
        except Exception as e:
            logger.error(f"Method 3 failed: {e}")
            return None

    def _debug_page_structure(self):
        logger.info("=" * 60)
        logger.info("DEBUG: Page Structure Analysis")
        logger.info("=" * 60)
        
        try:
            screenshot_path = "debug_screenshot.png"
            self.page.screenshot(path=screenshot_path, full_page=True)
            logger.info(f"Screenshot saved to: {screenshot_path}")
        except Exception as e:
            logger.error(f"Failed to save screenshot: {e}")
        
        try:
            current_url = self.page.url
            logger.info(f"Current URL: {current_url}")
        except Exception as e:
            logger.error(f"Failed to get URL: {e}")
        
        try:
            title = self.page.title()
            logger.info(f"Page Title: {title}")
        except Exception as e:
            logger.error(f"Failed to get title: {e}")
        
        try:
            body_text = self.page.inner_text("body")
            logger.info(f"Body text length: {len(body_text)} characters")
            logger.info(f"Body text preview: {body_text[:500]}")
        except Exception as e:
            logger.error(f"Failed to get body text: {e}")
        
        try:
            js_script = """
            var result = {
                allElements: document.querySelectorAll('*').length,
                divs: document.querySelectorAll('div').length,
                spans: document.querySelectorAll('span').length,
                canvases: document.querySelectorAll('canvas').length,
                tables: document.querySelectorAll('table').length,
                grids: document.querySelectorAll('[role=\"grid\"]').length,
                rows: document.querySelectorAll('[role=\"row\"]').length,
                cells: document.querySelectorAll('[role=\"gridcell\"]').length,
                titles: [],
                wpsClasses: []
            };
            
            var titleElements = document.querySelectorAll('[title]');
            for (var i = 0; i < Math.min(titleElements.length, 50); i++) {
                var title = titleElements[i].getAttribute('title');
                if (title && title.length < 100) {
                    result.titles.push(title);
                }
            }
            
            var allElements = document.querySelectorAll('*');
            var classSet = new Set();
            for (var i = 0; i < allElements.length; i++) {
                var cls = allElements[i].className;
                if (cls && cls.length > 0) {
                    cls.split(' ').forEach(function(c) {
                        if (c && c.length > 0) {
                            classSet.add(c);
                        }
                    });
                }
            }
            
            classSet.forEach(function(c) {
                if (c.includes('wps') || c.includes('sheet') || c.includes('table') || c.includes('grid') || c.includes('cell') || c.includes('row')) {
                    result.wpsClasses.push(c);
                }
            });
            
            return result;
            """
            
            result = self.page.evaluate(js_script)
            logger.info(f"Elements count:")
            logger.info(f"  - All: {result['allElements']}")
            logger.info(f"  - Divs: {result['divs']}")
            logger.info(f"  - Spans: {result['spans']}")
            logger.info(f"  - Canvases: {result['canvases']}")
            logger.info(f"  - Tables: {result['tables']}")
            logger.info(f"  - Grids: {result['grids']}")
            logger.info(f"  - Rows: {result['rows']}")
            logger.info(f"  - Cells: {result['cells']}")
            
            if result['titles']:
                logger.info(f"Titles found (first 20):")
                for title in result['titles'][:20]:
                    logger.info(f"  - {title}")
            
            if result['wpsClasses']:
                logger.info(f"WPS-related classes:")
                for cls in result['wpsClasses']:
                    logger.info(f"  - {cls}")
            
        except Exception as e:
            logger.error(f"JavaScript analysis failed: {e}")
        
        try:
            logger.info("\n--- JavaScript Global Variables ---")
            js_vars = """
            var interestingVars = [];
            var keywords = ['data', 'sheet', 'table', 'grid', 'document', 'workbook', 'cell', 'row', 'column', 'wps', 'kdocs'];
            
            for (var key in window) {
                try {
                    var lowerKey = key.toLowerCase();
                    if (keywords.some(function(k) { return lowerKey.includes(k); })) {
                        var val = window[key];
                        var type = typeof val;
                        var info = type;
                        if (type === 'object') {
                            info += ' (keys: ' + (val ? Object.keys(val).slice(0, 5).join(', ') : 'null') + ')';
                        } else if (type === 'string') {
                            info += ' (length: ' + val.length + ')';
                        }
                        interestingVars.push({ name: key, type: info });
                    }
                } catch(e) {
                    // ignore
                }
            }
            
            return interestingVars;
            """
            
            vars_result = self.page.evaluate(js_vars)
            if vars_result:
                logger.info(f"Found {len(vars_result)} interesting global variables:")
                for var in vars_result[:20]:
                    logger.info(f"  - {var['name']}: {var['type']}")
            
            logger.info("\n--- Checking window.wps and window.KDOCS ---")
            check_wps = """
            var result = {};
            if (window.wps) {
                result.wps = Object.keys(window.wps).slice(0, 20);
            }
            if (window.KDOCS) {
                result.KDOCS = Object.keys(window.KDOCS).slice(0, 20);
            }
            if (window.kdocs) {
                result.kdocs = Object.keys(window.kdocs).slice(0, 20);
            }
            if (window.sheet) {
                result.sheet = Object.keys(window.sheet).slice(0, 20);
            }
            return result;
            """
            
            wps_result = self.page.evaluate(check_wps)
            for key, value in wps_result.items():
                logger.info(f"  window.{key}: {value}")
            
        except Exception as e:
            logger.error(f"JavaScript variables check failed: {e}")
        
        logger.info("=" * 60)

    def extract_sheet_data(self, sheet_name=None):
        logger.info(f"Extracting sheet data{' for sheet: ' + sheet_name if sheet_name else ''}")
        try:
            if sheet_name:
                self._list_all_sheets()
                
                found = self._select_sheet(sheet_name)
                
                if found:
                    logger.info(f"Successfully switched to sheet: {sheet_name}")
                else:
                    logger.warning(f"Sheet '{sheet_name}' not found, using current sheet")
            
            return self.extract_table_data()
        except Exception as e:
            logger.error(f"Failed to extract sheet data: {e}")
            logger.error(traceback.format_exc())
            return []

    def _list_all_sheets(self):
        logger.info("Listing all available sheets on page...")
        
        sheet_selectors = [
            "div[title]",
            "span[title]",
            "button[title]",
            "[class*='sheet']",
            "[class*='tab']",
            "[class*='pane']"
        ]
        
        all_titles = set()
        
        for selector in sheet_selectors:
            try:
                elements = self.page.query_selector_all(selector)
                for element in elements:
                    try:
                        title = element.get_attribute("title")
                        if title and len(title) > 0 and len(title) < 50:
                            all_titles.add(title)
                    except:
                        pass
            except:
                pass
        
        js_script = """
        var titles = [];
        var elements = document.querySelectorAll('[title], [aria-label], [data-title]');
        for (var i = 0; i < elements.length; i++) {
            var title = elements[i].getAttribute('title') || elements[i].getAttribute('aria-label') || elements[i].getAttribute('data-title');
            if (title && title.length > 0 && title.length < 50) {
                titles.push(title);
            }
        }
        return titles;
        """
        
        try:
            js_titles = self.page.evaluate(js_script)
            if js_titles:
                for title in js_titles:
                    if title and len(title) < 50:
                        all_titles.add(title)
        except:
            pass
        
        if all_titles:
            logger.info(f"Found {len(all_titles)} potential sheet names:")
            for title in sorted(all_titles):
                logger.info(f"  - {title}")
        else:
            logger.warning("No sheet names found on page")

    def _select_sheet(self, sheet_name):
        logger.info(f"Trying to select sheet: '{sheet_name}'")
        
        selectors = [
            f"div[title='{sheet_name}']",
            f"span[title='{sheet_name}']",
            f"button[title='{sheet_name}']",
            f"div[title*='{sheet_name}']",
            f"span[title*='{sheet_name}']",
            f"button[title*='{sheet_name}']",
            f"[class*='sheet'][title*='{sheet_name}']",
            f"[class*='tab'][title*='{sheet_name}']",
            f"[class*='pane'][title*='{sheet_name}']",
            f".sheet-tab[title*='{sheet_name}']",
            f".wps-sheet-tab[title*='{sheet_name}']",
            f".wps-tab[title*='{sheet_name}']",
            f"[role='tab'][title*='{sheet_name}']",
            f"[role='tablist'] [title*='{sheet_name}']",
            f".wps-spreadsheet-sheet[title*='{sheet_name}']"
        ]
        
        for i, selector in enumerate(selectors):
            try:
                element = self.page.query_selector(selector)
                if element:
                    logger.info(f"✓ Found sheet '{sheet_name}' using selector #{i+1}: {selector}")
                    
                    try:
                        element.click()
                        logger.info(f"✓ Clicked on sheet element")
                    except:
                        logger.info(f"✓ Trying JavaScript click...")
                        self.page.evaluate(f"document.querySelector('{selector}').click()")
                    
                    time.sleep(3)
                    
                    current_title = element.get_attribute("title")
                    logger.info(f"✓ Current sheet title after click: '{current_title}'")
                    return True
            except Exception as e:
                logger.debug(f"Selector #{i+1} '{selector}' failed: {e}")
        
        logger.info(f"Trying JavaScript to find and click sheet '{sheet_name}'...")
        js_click_script = f"""
        var elements = document.querySelectorAll('[title*="{sheet_name}"], [title="{sheet_name}"]');
        for (var i = 0; i < elements.length; i++) {{
            if (elements[i].getAttribute('title').toLowerCase().includes('{sheet_name.toLowerCase()}')) {{
                elements[i].click();
                return 'Found and clicked via JS';
            }}
        }}
        return 'Not found via JS';
        """
        
        try:
            result = self.page.evaluate(js_click_script)
            logger.info(f"JavaScript click result: {result}")
            if result == 'Found and clicked via JS':
                time.sleep(3)
                return True
        except Exception as e:
            logger.error(f"JavaScript click failed: {e}")
        
        logger.warning(f"Failed to find sheet '{sheet_name}' with all selectors")
        return False

    def extract_text_by_keyword(self, keyword, before_chars=0, after_chars=100):
        logger.info(f"Extracting text by keyword: {keyword}")
        try:
            page_text = self.page.inner_text("body")
            pattern = rf".{{{before_chars}}}{re.escape(keyword)}.{{{after_chars}}}"
            matches = re.findall(pattern, page_text, re.DOTALL)
            if matches:
                logger.info(f"Found {len(matches)} matches")
                return matches
            return []
        except Exception as e:
            logger.error(f"Failed to extract text: {e}")
            logger.error(traceback.format_exc())
            return []

    def extract_text_by_selector(self, css_selector):
        logger.info(f"Extracting text by selector: {css_selector}")
        try:
            elements = self.page.query_selector_all(css_selector)
            texts = [element.inner_text() for element in elements]
            logger.info(f"Found {len(texts)} elements")
            return texts
        except Exception as e:
            logger.error(f"Failed to extract by selector: {e}")
            logger.error(traceback.format_exc())
            return []

    def extract_all_text(self):
        logger.info("Extracting all text")
        try:
            return self.page.inner_text("body")
        except Exception as e:
            logger.error(f"Failed to extract all text: {e}")
            logger.error(traceback.format_exc())
            return ""

    def close(self):
        logger.info("Closing browser...")
        
        try:
            self._save_storage_state()
        except Exception as e:
            logger.warning(f"Failed to save storage state during close: {e}")
        
        if self.page:
            try:
                self.page.close()
                logger.info("Page closed")
            except Exception as e:
                logger.warning(f"Failed to close page: {e}")
        
        if self.context:
            try:
                self.context.close()
                logger.info("Context closed")
            except Exception as e:
                logger.warning(f"Failed to close context: {e}")
        
        if self.browser:
            try:
                self.browser.close()
                logger.info("Browser closed")
            except Exception as e:
                logger.warning(f"Failed to close browser: {e}")
        
        if self.playwright:
            try:
                self.playwright.stop()
                logger.info("Playwright stopped")
            except Exception as e:
                logger.warning(f"Failed to stop playwright: {e}")
        
        logger.info("=" * 50)

    @staticmethod
    def extract_from_document(document_url, extraction_type="all", **kwargs):
        extractor = WPSExtractor(headless=False)
        result = None
        
        try:
            logger.info(f"Starting extraction process for: {document_url}")
            logger.info(f"Extraction type: {extraction_type}")
            
            extractor.open_wps_document(document_url)
            
            logger.info("Waiting for login...")
            if not extractor.wait_for_login():
                logger.error("Login failed or timeout")
                return None
            
            logger.info("Login successful, waiting for page to fully load...")
            time.sleep(3)
            
            logger.info("Starting data extraction...")
            
            if extraction_type == "keyword":
                keyword = kwargs.get("keyword", "")
                before_chars = kwargs.get("before_chars", 0)
                after_chars = kwargs.get("after_chars", 100)
                result = extractor.extract_text_by_keyword(keyword, before_chars, after_chars)
            elif extraction_type == "table":
                sheet_name = kwargs.get("sheet_name", None)
                if sheet_name:
                    result = extractor.extract_sheet_data(sheet_name)
                else:
                    result = extractor.extract_table_data()
            elif extraction_type == "selector":
                selector = kwargs.get("selector", "")
                result = extractor.extract_text_by_selector(selector)
            else:
                result = extractor.extract_all_text()
            
            if result:
                if isinstance(result, list):
                    total_rows = sum(len(t) for t in result) if isinstance(result[0], list) else len(result)
                    logger.info(f"Extraction successful: {len(result)} items, {total_rows} total rows")
                else:
                    logger.info(f"Extraction successful: {len(str(result))} characters")
                return result
            else:
                logger.warning("Extraction returned empty result, trying download method...")
        
        except Exception as e:
            logger.error(f"Browser extraction failed with exception: {e}")
            logger.error(traceback.format_exc())
        
        finally:
            logger.info("Closing extractor...")
            extractor.close()
        
        logger.info("=" * 60)
        logger.info("FALLBACK: Trying download method")
        logger.info("=" * 60)
        
        sheet_name = kwargs.get("sheet_name", None)
        download_result = WPSExtractor.extract_via_download(document_url, sheet_name)
        
        if download_result:
            logger.info("Download extraction succeeded!")
            return download_result
        
        logger.error("All extraction methods failed")
        return None