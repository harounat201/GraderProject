import json
import os
import shutil
import tempfile
import zipfile
from utils.xml_helper import parse_workbook_to_json
from grader import grade_submission

def verify_deep_charts():
    # 1. Verify JSON extraction directly
    sub_path = "Data Visualization (2) copy.xlsm"
    print(f"--- 1. Verifying JSON extraction for: {sub_path} ---")
    
    td = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(sub_path, 'r') as z:
            z.extractall(td)
        
        data = parse_workbook_to_json(td)
        
        # Look for drawings with details
        found_chart = False
        for sheet_name, sheet_data in data.get("sheets", {}).items():
            drawings = sheet_data.get("metadata", {}).get("drawings", [])
            for d in drawings:
                if isinstance(d, dict) and d.get("type") == "chart":
                    details = d.get("details", {})
                    if "type" in details and "axes" in details:
                         print(f"SUCCESS: Found chart in sheet '{sheet_name}'")
                         print(f"Chart Type: {details['type']}")
                         print(f"Axes: {list(details['axes'].keys())}")
                         found_chart = True
                         break
            if found_chart: break
            
        if not found_chart:
            print("FAIL: No chart details extracted in JSON.")
            return

    finally:
        shutil.rmtree(td)

    # 2. Verify AI Grading (Full Pipeline)
    print(f"\n--- 2. Verifying AI Grading for: {sub_path} ---")
    results = grade_submission(sub_path)
    report = results.get('report', {})
    
    print(f"Final Score: {report.get('score', {}).get('earned')}/{report.get('score', {}).get('max')}")
    print(f"Summary Snippet: {report.get('summary')[:200]}...")
    
    # Check if specific criteria were met (Task 1 for axis scale, Task 6 for Scatter)
    criteria = report.get('criteria', [])
    for c in criteria:
        name = c.get('name', '')
        earned = c.get('earned', 0)
        max_v = c.get('max', 1)
        if "Task 1" in name or "Task 6" in name:
            print(f"Criterion '{name}': {earned}/{max_v}")

if __name__ == "__main__":
    verify_deep_charts()
