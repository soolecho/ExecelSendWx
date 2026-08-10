import logging

logger = logging.getLogger(__name__)


class FilterCondition:
    def __init__(self, column_name, operator, value):
        self.column_name = column_name
        self.operator = operator
        self.value = value

    def __repr__(self):
        return f"FilterCondition(column='{self.column_name}', operator='{self.operator}', value='{self.value}')"


class TableProcessor:
    def __init__(self, table_data):
        self.table_data = table_data
        self.headers = []
        self.rows = []
        if table_data and isinstance(table_data, list) and len(table_data) > 0:
            self._parse_table(table_data)

    def _parse_table(self, table):
        if not table or len(table) == 0:
            return
        
        self.headers = table[0]
        self.rows = table[1:]
        
        logger.info(f"Parsed table with {len(self.headers)} columns and {len(self.rows)} rows")

    def get_headers(self):
        return self.headers

    def get_header_index(self, header_name):
        for i, header in enumerate(self.headers):
            if header_name in header or header in header_name:
                return i
        return -1

    def _match_condition(self, row, condition):
        col_index = self.get_header_index(condition.column_name)
        if col_index < 0:
            return False
        
        if col_index >= len(row):
            cell_value = ""
        else:
            cell_value = str(row[col_index])
        
        operator = condition.operator
        value = str(condition.value)
        
        if operator == "equals":
            return cell_value == value
        elif operator == "contains":
            return value in cell_value
        elif operator == "not_contains":
            return value not in cell_value
        elif operator == "empty":
            return not cell_value.strip()
        elif operator == "not_empty":
            return bool(cell_value.strip())
        elif operator == "greater_than":
            try:
                return float(cell_value) > float(value)
            except:
                return False
        elif operator == "less_than":
            try:
                return float(cell_value) < float(value)
            except:
                return False
        elif operator == "equals_or_greater":
            try:
                return float(cell_value) >= float(value)
            except:
                return False
        elif operator == "equals_or_less":
            try:
                return float(cell_value) <= float(value)
            except:
                return False
        elif operator == "date_equals":
            return cell_value.strip() == value.strip()
        elif operator == "date_after":
            try:
                from datetime import datetime
                cell_date = datetime.strptime(cell_value.strip(), "%Y-%m-%d")
                target_date = datetime.strptime(value.strip(), "%Y-%m-%d")
                return cell_date > target_date
            except:
                return False
        elif operator == "date_before":
            try:
                from datetime import datetime
                cell_date = datetime.strptime(cell_value.strip(), "%Y-%m-%d")
                target_date = datetime.strptime(value.strip(), "%Y-%m-%d")
                return cell_date < target_date
            except:
                return False
        else:
            return True

    def filter_by_conditions(self, conditions):
        if not conditions:
            return self.rows[:]
        
        filtered_rows = []
        for row in self.rows:
            match_all = True
            for condition in conditions:
                if not self._match_condition(row, condition):
                    match_all = False
                    break
            if match_all:
                filtered_rows.append(row)
        
        logger.info(f"Filtered {len(filtered_rows)} rows with {len(conditions)} conditions")
        return filtered_rows

    def filter_by_name(self, name, name_column_index=None, name_column_name=None):
        if name_column_index is None and name_column_name is not None:
            name_column_index = self.get_header_index(name_column_name)
        
        if name_column_index < 0:
            logger.warning(f"Name column not found: {name_column_name}")
            return []
        
        filtered_rows = []
        for row in self.rows:
            if len(row) > name_column_index and name in row[name_column_index]:
                filtered_rows.append(row)
        
        logger.info(f"Filtered {len(filtered_rows)} rows for name: {name}")
        return filtered_rows

    def extract_columns(self, rows, column_indices=None, column_names=None):
        if column_indices is None and column_names is not None:
            column_indices = [self.get_header_index(name) for name in column_names]
        
        column_indices = [i for i in column_indices if i >= 0]
        
        if not column_indices:
            logger.warning("No valid columns specified")
            return []
        
        extracted_data = []
        for row in rows:
            extracted_row = [row[i] if i < len(row) else "" for i in column_indices]
            extracted_data.append(extracted_row)
        
        logger.info(f"Extracted {len(extracted_data)} rows with {len(column_indices)} columns")
        return extracted_data

    def format_row_data(self, row, column_names=None, column_indices=None):
        if column_names is None:
            if column_indices is None:
                column_names = self.headers
            else:
                column_names = [self.headers[i] if i < len(self.headers) else f"列{i+1}" for i in column_indices]
        
        result = []
        for i, (name, value) in enumerate(zip(column_names, row)):
            if column_indices is not None and i < len(column_indices):
                result.append(f"{name}: {value}")
            else:
                result.append(f"{name}: {value}")
        
        return "\n".join(result)

    def get_person_table_data(self, name, name_column, extract_columns, conditions=None):
        name_index = self.get_header_index(name_column)
        if name_index < 0:
            logger.error(f"Name column '{name_column}' not found in headers: {self.headers}")
            return None
        
        extract_indices = [self.get_header_index(col) for col in extract_columns]
        extract_indices = [i for i in extract_indices if i >= 0]
        
        if not extract_indices:
            logger.error(f"No valid extract columns found: {extract_columns}")
            return None
        
        if conditions:
            filtered_rows = self.filter_by_conditions(conditions)
            name_condition = FilterCondition(name_column, "contains", name)
            filtered_rows = [row for row in filtered_rows if self._match_condition(row, name_condition)]
        else:
            filtered_rows = self.filter_by_name(name, name_column_index=name_index)
        
        if not filtered_rows:
            logger.warning(f"No data found for person: {name}")
            return None
        
        extracted_data = self.extract_columns(filtered_rows, column_indices=extract_indices)
        
        extract_names = [self.headers[i] for i in extract_indices]
        return {
            "headers": extract_names,
            "rows": extracted_data
        }

    def get_person_data(self, name, name_column, extract_columns, conditions=None):
        table_data = self.get_person_table_data(
            name,
            name_column,
            extract_columns,
            conditions
        )
        if not table_data:
            return None
        
        return self.format_table_data(table_data)

    def format_table_data(self, table_data):
        formatted_results = []
        for row in table_data["rows"]:
            formatted = self.format_row_data(
                row,
                column_names=table_data["headers"]
            )
            formatted_results.append(formatted)
        
        return "\n\n".join(formatted_results)

    def get_all_persons(self, name_column, conditions=None):
        name_index = self.get_header_index(name_column)
        if name_index < 0:
            logger.error(f"Name column '{name_column}' not found")
            return []
        
        if conditions:
            filtered_rows = self.filter_by_conditions(conditions)
        else:
            filtered_rows = self.rows
        
        persons = set()
        for row in filtered_rows:
            if len(row) > name_index and row[name_index]:
                persons.add(row[name_index])
        
        logger.info(f"Found {len(persons)} unique persons with filters")
        return sorted(list(persons))

    def get_person_to_wechat_mapping(self, name_column, wechat_column):
        name_index = self.get_header_index(name_column)
        wechat_index = self.get_header_index(wechat_column)
        
        if name_index < 0 or wechat_index < 0:
            logger.error(f"Columns not found - name: {name_column}, wechat: {wechat_column}")
            return {}
        
        mapping = {}
        for row in self.rows:
            if len(row) > max(name_index, wechat_index):
                name = row[name_index]
                wechat = row[wechat_index]
                if name and wechat:
                    mapping[name] = wechat
        
        logger.info(f"Created mapping for {len(mapping)} persons")
        return mapping
