import os
import json
import openpyxl

def extract_serial_numbers_from_excel(excel_file_path):
    """Extract serial numbers from the provided Excel file."""
    try:
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        ws = wb.active
        serial_numbers = set()
        for row in ws.iter_rows(min_row=3, min_col=1, max_col=1, values_only=True):
            cell_value = str(row[0]).strip().upper() if row[0] else ''
            if cell_value.isdigit() and len(cell_value) <= 8:
                serial_numbers.add(cell_value)
                print(f"Extracted serial number: {cell_value} from Excel sheet")
        return serial_numbers
    except Exception as e:
        print(f"Error reading from Excel: {e}")
        return set()

def check_serial_number_folders(base_path, serial_numbers):
    """Check for folders matching the serial numbers extracted from Excel within customer folders."""
    found_serial_numbers = {}
    errors = {}
    for letter in os.listdir(base_path):
        letter_path = os.path.join(base_path, letter)
        if os.path.isdir(letter_path):
            for customer_folder in os.listdir(letter_path):
                customer_path = os.path.join(letter_path, customer_folder)
                if os.path.isdir(customer_path):
                    # Check each item inside the customer folder
                    for item in os.listdir(customer_path):
                        full_path = os.path.join(customer_path, item)
                        # Split the item name to handle cases where the serial number is followed by additional text
                        serial_candidate = item.split()[0]  # Assuming the serial number is always the first part before a space
                        if serial_candidate.isdigit() and serial_candidate in serial_numbers:
                            found_serial_numbers[serial_candidate] = {
                                "serialnumber": serial_candidate,
                                "serialfullpath": full_path,
                                "serialpath": full_path.split("\\")
                            }
                            print(f"Found and logged serial number: {serial_candidate} at {full_path}")
                        elif serial_candidate.isdigit() and len(serial_candidate) <= 8 and serial_candidate not in serial_numbers:
                            errors.setdefault("SERIALNUMBERNOTINSPREADSHEET", []).append(serial_candidate)

    return found_serial_numbers, errors

def main():
    basepath = 'P:/CONREC/CUSTOMERS'
    excel_file_path = "P:/CONREC/Serial Number List.xlsx"

    print("\nParsing all CONREC Files in KONTEK's Network...\n")
    serial_numbers = extract_serial_numbers_from_excel(excel_file_path)
    serialnums, errors = check_serial_number_folders(basepath, serial_numbers)

    with open("serialnum.json", "w") as f:
        json.dump(serialnums, f, indent=4)
    with open("errors.json", "w") as f:
        json.dump(errors, f, indent=4)

    print("\nParsing Complete!\n")
    print(f"Logged {len(serialnums)} found serial number folders to serialnum.json")
    print(f"Logged {len(errors.get('SERIALNUMBERNOTINSPREADSHEET', []))} serial numbers not in spreadsheet to errors.json")

if __name__ == "__main__":
    main()
