def split_excel_by_row(excel_data, sheet_name):
    def validate_excel_data(data):
        if not data:
            raise ValueError("Excel data is empty.")

    def extract_rows(data):
        return [row for row in data if row]

    def process_rows(rows):
        # Implement your row processing logic here
        return rows

    validate_excel_data(excel_data)
    rows = extract_rows(excel_data)
    processed_rows = process_rows(rows)
    return processed_rows
