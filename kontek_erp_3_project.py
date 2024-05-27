import os
import json
import openpyxl

basepath = 'P:/CONREC/ARCHIVE'
projects = {}
errors = {}

# Extracts project numbers, including those with suffixes from the Excel file

def extract_project_numbers_from_excel(excel_file_path):
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

# Checks through the network folders to find project folders too, and logs them if they match the project numbers found in the Excel file.
    
def check_project_folder(base_path):
    try:
        for root, dirs, files in os.walk(base_path, topdown=True):
            for folder in dirs:
                full_path = os.path.join(root, folder)
                folder_parts = folder.split()
                for part in folder_parts:
                    if part.startswith('K') and len(part) >= 8 and part[1:8].isdigit():
                        if '-' in part:
                            project_base, suffix = part.split('-', 1)
                            if suffix.isalpha() and 1 <= len(suffix) <= 3:
                                project_number = project_base + '-' + suffix
                            else:
                                project_number = project_base
                        else:
                            project_number = part[:8]  
                        
                        if project_number not in projects:
                            projects[project_number] = {
                                "projectnumber": project_number,
                                "projectfullpath": full_path,
                                "projectpath": full_path.split("\\")
                            }
                            print(f"Found and logged project: {project_number} at {full_path}")
                            break  
    except Exception as e:
        print(f"Error checking project folder: {e}")

# Finds unmatched projects in both the Excel file and network and prints them out for the user to fix

def find_unmatched_projects(excel_project_numbers):
    try:
        found_projects = set(projects.keys())
        missing_projects = excel_project_numbers.difference(found_projects)
        if missing_projects:
            errors["PROJECTNUMBERSFOLDERNOTFOUND"] = list(missing_projects)
            for mp in missing_projects:
                print(f"Missing project number: {mp} not found in directories")
    except Exception as e:
        print(f"Error finding unmatched projects: {e}")

# Print comments for clarity, but can be omitted

def main():
    print("\nParsing all CONREC Files in KONTEK's Network...\n")
    excel_file_path = "P:/CONREC/Serial Number List.xlsx"
    excel_project_numbers = extract_project_numbers_from_excel(excel_file_path)
    check_project_folder(basepath)
    find_unmatched_projects(excel_project_numbers)

# Creating projects.json and errors.json files to store parsed data

    with open("projects.json", "w") as f:
        json.dump(projects, f, indent=4)
    with open("errors.json", "w") as f:
        json.dump(errors, f, indent=4)

    print("\nParsing Complete!\n")
    print(f"Logged {len(projects)} found projects to projects.json")
    print(f"Logged {len(errors.get('PROJECTNUMBERSFOLDERNOTFOUND', []))} missing projects to errors.json")

if __name__ == "__main__":
    main()