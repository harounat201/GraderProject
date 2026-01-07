import json
import os
from grader import grade_submission

def verify_false_positive_fix():
    # Use the Data Visualization template file
    sub_path = "Data Visualization (2) copy.xlsm"
    
    # We'll use the embedded rubric (which the grader auto-detects)
    # or pass it explicitly if needed. In this case, grade_submission
    # will find the 'Scoring Guide' sheet.
    
    print(f"--- Verifying fix with template: {sub_path} ---")
    
    if not os.path.exists(sub_path):
        print(f"Error: {sub_path} not found.")
        return
        
    try:
        results = grade_submission(sub_path)
        
        report = results.get('report', {})
        score = report.get('score', {}).get('earned', -1)
        max_score = report.get('score', {}).get('max', -1)
        
        print(f"Score: {score}/{max_score}")
        print(f"Summary: {report.get('summary')}")
        
        if score == 0:
            print("PASS: The template correctly received a 0.")
        elif score < 5:
            print(f"PARTIAL PASS: The template received a very low score ({score}).")
        else:
            print(f"FAIL: The template still received a high score ({score}).")
            
    except Exception as e:
        print(f"Execution Error: {str(e)}")

if __name__ == "__main__":
    verify_false_positive_fix()
