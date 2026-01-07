import json
import os
import shutil
import tempfile
import zipfile
from utils.xml_helper import parse_workbook_to_json

def verify_cell_styles():
    sub_path = "grossjordan_34944_1593395_Data Visualization - Jordan Gross.xlsm"
    print(f"--- Verifying cell styles for: {sub_path} ---")
    
    td = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(sub_path, 'r') as z:
            z.extractall(td)
        
        data = parse_workbook_to_json(td)
        
        # Check a few cells in Task 8 sheet
        sheet_found = False
        for sheet_name, sheet_data in data.get("sheets", {}).items():
            if "Task 8" in sheet_name or "T 8" in sheet_name:
                print(f"Checking sheet: {sheet_name}")
                cells = sheet_data.get("cells", {})
                
                # Print styles for first 5 non-empty cells
                count = 0
                for coord, cell_data in cells.items():
                    if 'style' in cell_data:
                        print(f"Cell {coord}: {json.dumps(cell_data['style'])}")
                        count += 1
                        if count >= 5: break
                
                if count == 0:
                    print("No non-default styles found in this sheet.")
                sheet_found = True
                break
        
        if not sheet_found:
             # Just check any sheet
             print("Task 8 sheet not found, checking first available sheet.")
             first_sheet = next(iter(data.get("sheets", {}).values()))
             cells = first_sheet.get("cells", {})
             count = 0
             for coord, cell_data in cells.items():
                 if 'style' in cell_data:
                     print(f"Cell {coord}: {json.dumps(cell_data['style'])}")
                     count += 1
                     if count >= 5: break

    finally:
        shutil.rmtree(td)

if __name__ == "__main__":
    verify_cell_styles()
