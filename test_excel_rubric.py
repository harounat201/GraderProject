import argparse
import json
import os
from grader import grade_submission

def test_excel_rubric():
    # We use DataManagement.xlsm as the RUBRIC file (since it has Scoring Guide)
    # And we use it as the SUBMISSION file.
    
    file_path = "DataManagement.xlsm"
    
    if not os.path.exists(file_path):
        print("Test file not found.")
        return
        
    print(f"--- Grading Submission using Rubric extracted from {file_path} ---")
    
    # We pass rubric_path as the xlsx file
    results = grade_submission(file_path, rubric_path=file_path)
    
    print(json.dumps(results, indent=2))
    
if __name__ == "__main__":
    test_excel_rubric()
