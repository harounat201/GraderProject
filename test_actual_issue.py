import json
import os
from grader import grade_submission

def test_user_files():
    sub_path = "Grade Automation for Graduate Coursework_ A Product Overview (1).docx"
    rubric_path = "Study Paper Rubric [TEST].xlsx"
    
    print(f"--- Testing with files: {sub_path} and {rubric_path} ---")
    
    if not os.path.exists(sub_path):
        print(f"Error: {sub_path} not found.")
        return
    if not os.path.exists(rubric_path):
        print(f"Error: {rubric_path} not found.")
        return
        
    try:
        results = grade_submission(sub_path, rubric_path=rubric_path)
        
        # Check structure
        report = results.get('report', {})
        prompt = results.get('prompt', '')
        
        print(f"Global Keys: {list(results.keys())}")
        print(f"Report Keys: {list(report.keys())}")
        print(f"Prompt length: {len(prompt)}")
        
        if "error" in report:
            print(f"Report Error: {report['error']}")
        
        # Save results for inspection
        with open("reproduction_results.json", "w") as f:
            json.dump(results, f, indent=2)
            
        print("Done. See reproduction_results.json")
        
    except Exception as e:
        print(f"Execution Error: {str(e)}")

if __name__ == "__main__":
    test_user_files()
