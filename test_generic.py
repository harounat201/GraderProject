import argparse
import json
import os
from grader import grade_submission

def test_generic():
    # Create dummy files
    sub_path = "submission_test.txt"
    rubric_path = "rubric_test.txt"
    
    with open(sub_path, 'w') as f:
        f.write("The quick brown fox jumps over the lazy dog. It was a sunny day.")
        
    with open(rubric_path, 'w') as f:
        f.write("Scoring Guide:\n1. Mention a fox (5 points)\n2. Mention a dog (5 points)\n3. Mention the weather (5 points)")
        
    print(f"--- Grading Text Submission using Text Rubric ---")
    
    results = grade_submission(sub_path, rubric_path=rubric_path)
    print(json.dumps(results, indent=2))
    
    # Cleanup
    if os.path.exists(sub_path): os.remove(sub_path)
    if os.path.exists(rubric_path): os.remove(rubric_path)

if __name__ == "__main__":
    test_generic()
