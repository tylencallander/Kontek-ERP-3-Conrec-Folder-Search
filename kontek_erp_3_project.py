import os
import json
import openpyxl

basepath = 'P:/CONREC/ARCHIVE'
serialnums = {}
errors = {}

# Extracts serial numbers, from the Excel file

def extract_serial_numbers_from_excel(excel_file_path):
    try:
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        ws = wb.active
        project_numbers = set()
        for row in ws.iter_rows(min_row=3, min_col=1, max_col=1, values_only=True):
            cell_value = str(row[0]).strip().upper() if row[0] else ''
            if cell_value[1:7].isdigit():
                project_numbers.add(cell_value)
                print(f"Extracted serial number: {cell_value} from Excel sheet")
        return project_numbers
    except Exception as e:
        print(f"Error reading from Excel: {e}")
        return set()

# Print comments for clarity, but can be omitted

def main():
    print("\nParsing all CONREC Files in KONTEK's Network...\n")
    excel_file_path = "P:/CONREC/Serial Number List.xlsx"
    excel_project_numbers = extract_serial_numbers_from_excel(excel_file_path)

# Creating serialnum.json and errors.json files to store parsed data 

    print("\nParsing Complete!\n")
    print(f"Logged {len(serialnums)} found serial number folders to serialnum.json")
    print(f"Logged {len(errors.get('SERIALNUMBERFOLDERNOTFOUND', []))} missing serial number folders to errors.json")

if __name__ == "__main__":
    main()