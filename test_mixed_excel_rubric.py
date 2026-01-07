import os
import json
from grader import grade_submission

def test_mixed_excel_rubric():
    print("--- Test: Text Submission + Excel Rubric ---")
    
    # 1. Create a dummy text submission
    sub_path = "essay.txt"
    with open(sub_path, 'w') as f:
        f.write("This is a demo essay. It talks about many things.")
        
    # 2. Use an existing Excel file as a rubric (e.g. DataManagement.xlsm)
    # Even if it's not a 'Scoring Guide', it should now fall back to raw data.
    rubric_path = "DataManagement.xlsm"
    
    results = grade_submission(sub_path, rubric_path=rubric_path)
    
    if os.path.exists(sub_path): os.remove(sub_path)

    # Check top-level wrapping
    print(f"Global Keys: {list(results.keys())}")
    
    report = results.get('report', {})
    prompt = results.get('prompt', '')
    
    print(f"Report Keys: {list(report.keys())}")
    print(f"Prompt length: {len(prompt)}")
    
    if "score" in report and len(prompt) > 100:
        print("PASS: Prompt captured and Report structure correct.")
    else:
        print("FAIL: Check results.")

    # Show snippet of report
    print(json.dumps(report, indent=2)[:500])

if __name__ == "__main__":
    test_mixed_excel_rubric()
