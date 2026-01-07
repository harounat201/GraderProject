import zipfile
import tempfile
import os
import shutil
import json
from utils.xml_helper import parse_workbook_to_json
from utils.rubric_extractor import extract_rubric_from_sheet

def inspect_user_rubric():
    path = "Study Paper Rubric [TEST].xlsx"
    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(path, 'r') as z:
            z.extractall(temp_dir)
        
        print("--- Testing structured extraction ---")
        extracted = extract_rubric_from_sheet(temp_dir)
        print(f"Extracted: {json.dumps(extracted, indent=2)}")
        
        print("\n--- Testing raw extraction ---")
        raw = parse_workbook_to_json(temp_dir)
        # Count keys in Sheet1 cells
        if "sheets" in raw and "Sheet1" in raw["sheets"]:
            print(f"Sheet1 cells: {len(raw['sheets']['Sheet1']['cells'])}")
            # Show first few cells
            cells = list(raw['sheets']['Sheet1']['cells'].items())[:5]
            print(f"Sample cells: {json.dumps(cells, indent=2)}")
        else:
            print(f"Sheets found: {list(raw.get('sheets', {}).keys())}")
        
    finally:
        shutil.rmtree(temp_dir)

if __name__ == "__main__":
    inspect_user_rubric()
