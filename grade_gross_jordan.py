import json
import os
from grader import grade_submission

def grade_gross_jordan():
    sub_path = "grossjordan_34944_1593395_Data Visualization - Jordan Gross.xlsm"
    
    print(f"--- Grading Student: Gross Jordan ---")
    print(f"File: {sub_path}")
    
    if not os.path.exists(sub_path):
        print(f"Error: {sub_path} not found.")
        return
        
    try:
        results = grade_submission(sub_path)
        report = results.get('report', {})
        
        print(f"Final Score: {report.get('score', {}).get('earned')}/{report.get('score', {}).get('max')}")
        print(f"Summary: {report.get('summary')}")
        
        print("\n--- Criteria Breakdown ---")
        for c in report.get('criteria', []):
            status = "✅" if c.get('earned') == c.get('max') else "❌" if c.get('earned') == 0 else "⚠️"
            print(f"{status} {c.get('name')}: {c.get('earned')}/{c.get('max')}")
            print(f"   Feedback: {c.get('feedback')}")
            
    except Exception as e:
        print(f"Execution Error: {str(e)}")

if __name__ == "__main__":
    grade_gross_jordan()
