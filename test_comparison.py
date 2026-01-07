import argparse
import json
import os
from grader import grade_submission

def test_comparison():
    # We will use DataManagement.xlsm as BOTH student and answer key
    # It should result in a perfect match.
    
    file_path = "DataManagement.xlsm"
    
    if not os.path.exists(file_path):
        print("Test file not found.")
        return
        
    print(f"Comparing {file_path} against itself...")
    results = grade_submission(file_path, answer_key_path=file_path)
    
    print(json.dumps(results, indent=2))
    
if __name__ == "__main__":
    test_comparison()
