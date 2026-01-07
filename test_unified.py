import argparse
import json
import os
from grader import grade_submission

def test_unified():
    file_path = "DataManagement.xlsm"
    
    if not os.path.exists(file_path):
        print("Test file not found.")
        return
        
    print(f"--- Grading {file_path} in HYBRID Mode (Embedded Rubric + Self as Key) ---")
    
    # We pass answer_key_path. Grader should also find the embedded rubric.
    # So it should trigger the Logic: rubric_data AND answer_key_data
    results = grade_submission(file_path, answer_key_path=file_path)
    
    print(json.dumps(results, indent=2))
    
if __name__ == "__main__":
    test_unified()
