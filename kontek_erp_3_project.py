import os
import json
import openpyxl

# Extracts project numbers from the Excel file

def extract_serial_numbers_from_excel(excel_file_path):
    try:
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        ws = wb.active
        serial_numbers = set()
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
            cell_value = str(row[0]).strip().upper() if row[0] else ''
            if cell_value.isdigit() and len(cell_value) <= 8:
                serial_numbers.add(cell_value)
                print(f"Extracted serial number: {cell_value} from Excel sheet")
        return serial_numbers
    except Exception as e:
        print(f"Error reading from Excel: {e}")
        return set()
    
# Checks through the network folders to find serial number folders and logs them if they match the serial numbers found in the Excel file

def check_serial_number_folders(base_path, serial_numbers):
    found_serial_numbers = {}
    errors = {}
    for letter in os.listdir(base_path):
        letter_path = os.path.join(base_path, letter)
        if os.path.isdir(letter_path):
            for customer_folder in os.listdir(letter_path):
                customer_path = os.path.join(letter_path, customer_folder)
                if os.path.isdir(customer_path):
                    for item in os.listdir(customer_path):
                        full_path = os.path.join(customer_path, item)
                        serial_candidate = item.split()[0] 
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
    basepath = 'P:/CONREC/CUSTOMERS' # ONLY DID CUSTOMERS FOLDER SERIAL NUMBER
    excel_file_path = "P:/CONREC/CONREC PROJECT SERIAL NUMBERS.xlsx"

    print("\nParsing all CONREC Files in KONTEK's Network...\n")
    serial_numbers = extract_serial_numbers_from_excel(excel_file_path)
    serialnums, errors = check_serial_number_folders(basepath, serial_numbers)

# Creating projects.json and errors.json files to store parsed data

    with open("serialnum.json", "w") as f:
        json.dump(serialnums, f, indent=4)
    with open("errors.json", "w") as f:
        json.dump(errors, f, indent=4)

    print("\nParsing Complete!\n")
    print(f"Logged {len(serialnums)} found serial number folders to serialnum.json")
    print(f"Logged {len(errors.get('SERIALNUMBERNOTINSPREADSHEET', []))} serial numbers not in spreadsheet to errors.json")

if __name__ == "__main__":
    main()