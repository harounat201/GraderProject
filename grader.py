import argparse
import json
import os
import zipfile
import tempfile
import shutil
from utils.xml_helper import get_sheet_map, get_shared_strings, parse_sheet_full
from utils.evaluator import evaluate_task
from utils.rubric_extractor import extract_rubric_from_sheet

def grade_submission(submission_path, rubric_data=None, rubric_path=None, answer_key_path=None):
    """
    Grades a submission. 
    If answer_key_path is provided, uses AI Comparison Grading.
    Otherwise uses Rubric (dict/path/embedded).
    """
    
    # Prepare vars
    temp_dir = tempfile.mkdtemp()
    temp_dir_key = None
    
    try:
        # Prepare Data for AI
        from utils.xml_helper import parse_workbook_to_json
        from utils.text_extractor import extract_text_from_file
        from utils.llm_helper import grade_student_work
        
        # 1. Parse Student Data
        student_data = None
        if submission_path.endswith(('.xlsx', '.xlsm')):
            # Zip processing for Excel
            # (temp_dir is already created above)
            with zipfile.ZipFile(submission_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            student_data = parse_workbook_to_json(temp_dir)
        elif submission_path.lower().endswith(('.docx', '.txt')):
            student_data = extract_text_from_file(submission_path)
            
        if not student_data:
            return {"error": "Unsupported submission format or empty file."}
        
        # 2. Parse Answer Key (Optional)
        answer_key_data = None
        if answer_key_path:
            if answer_key_path.endswith(('.xlsx', '.xlsm')):
                temp_dir_key = tempfile.mkdtemp()
                with zipfile.ZipFile(answer_key_path, 'r') as z_key:
                    z_key.extractall(temp_dir_key)
                answer_key_data = parse_workbook_to_json(temp_dir_key)
            elif answer_key_path.lower().endswith(('.docx', '.txt')):
                answer_key_data = extract_text_from_file(answer_key_path)

        # 3. Load Rubric (Optional)
        rubric = None
        rubric_source = "none"
        
        if rubric_data:
            rubric = rubric_data
            rubric_source = "provided"
        elif rubric_path:
            if rubric_path.endswith('.json'):
                try:
                    with open(rubric_path, 'r') as f:
                        rubric = json.load(f)
                    rubric_source = "provided"
                except:
                    pass
            elif rubric_path.lower().endswith(('.xlsx', '.xlsm')):
                # Extract rubric from the provided Excel file
                temp_rubric_dir = tempfile.mkdtemp()
                try:
                    with zipfile.ZipFile(rubric_path, 'r') as z_rubric:
                        z_rubric.extractall(temp_rubric_dir)
                    extracted_rubric = extract_rubric_from_sheet(temp_rubric_dir)
                    if extracted_rubric:
                        rubric = extracted_rubric
                        rubric_source = "provided_excel"
                    else:
                        # Fallback: Parse entire workbook as raw data
                        rubric = parse_workbook_to_json(temp_rubric_dir)
                        rubric_source = "provided_excel_raw"
                except Exception as e:
                     print(f"Excel Rubric Error: {e}")
                finally:
                    if os.path.exists(temp_rubric_dir):
                        shutil.rmtree(temp_rubric_dir)
            elif rubric_path.lower().endswith(('.docx', '.txt')):
                 # Text Rubric
                 rubric = extract_text_from_file(rubric_path)
                 rubric_source = "provided_text"

        # If no explicit rubric and using Excel, check for embedded
        if not rubric and isinstance(student_data, dict):
            # Only check embedded if it's an Excel submission (which we have unzip dir for)
            if os.path.exists(temp_dir):
                 extracted = extract_rubric_from_sheet(temp_dir)
                 if extracted:
                     rubric = extracted
                     rubric_source = "embedded"

        # 4. Execute AI Grading
        # If we have NEITHER rubric nor key, we can't grade.
        if not rubric and not answer_key_data:
             if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
             return {
                 "report": {"error": "Cannot grade: No Rubric provided/found AND no Answer Key provided."},
                 "prompt": "No prompt generated (missing context)."
             }
             
        ai_response = grade_student_work(student_data, rubric_data=rubric, answer_key_data=answer_key_data)
        
        # Ensure ai_response is properly structured
        if "report" not in ai_response:
             # Fallback if somehow return was raw (e.g. from an old version)
             ai_response = {"report": ai_response, "prompt": "Prompt not captured."}

        # Inject metadata into the REPORT so UI sees it
        ai_response['report']['rubric_source'] = rubric_source
        ai_response['report']['mode'] = "ai_unified"
        
        return ai_response

    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        if temp_dir_key and os.path.exists(temp_dir_key):
            shutil.rmtree(temp_dir_key)

def prepare_grading_context(uploaded_files):
    """
    Extracts text from various file formats and combines them into a single context string.
    """
    from utils.xml_helper import parse_workbook_to_json
    from utils.text_extractor import extract_text_from_file
    
    contexts = []
    
    for f in uploaded_files:
        # Create a temp file path to read from
        suffix = f".{f.name.split('.')[-1]}"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(f.getvalue() if hasattr(f, 'getvalue') else open(f, 'rb').read())
            tmp_path = tmp.name
            
        try:
            if tmp_path.endswith(('.xlsx', '.xlsm')):
                td = tempfile.mkdtemp()
                with zipfile.ZipFile(tmp_path, 'r') as z:
                    z.extractall(td)
                data = parse_workbook_to_json(td)
                contexts.append(f"FILE: {f.name}\nCONTENT (JSON):\n{json.dumps(data, indent=2)}")
                shutil.rmtree(td)
            else:
                text = extract_text_from_file(tmp_path)
                contexts.append(f"FILE: {f.name}\nCONTENT:\n{text}")
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
                
    return "\n\n---\n\n".join(contexts)

def main():
    parser = argparse.ArgumentParser(description="Excel XML Grader")
    parser.add_argument("--submission", required=True, help="Path to .xlsx file")
    parser.add_argument("--rubric", help="Path to rubric.json (Optional if embedded in file)")
    parser.add_argument("--context", help="Path to assignment context (optional)")
    
    args = parser.parse_args()
    
    results = grade_submission(args.submission, rubric_path=args.rubric)
    print(json.dumps(results, indent=2))

if __name__ == "__main__":
    main()
