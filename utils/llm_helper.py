import re
from google import genai
import os
import json

# Configure API Key securely via environment variable
api_key = os.environ.get('GEMINI_API_KEY')
if not api_key:
    # Use a placeholder or raise an error in a production environment
    # For Streamlit, this will prompt the user to configure secrets
    api_key = "MISSING_API_KEY"

client = genai.Client(api_key=api_key)

ATOMIC_RUBRIC_PROMPT = """
You are an expert Education Consultant specializing in Technical & Data Assessment.
Your task is to generate a highly granular, ATOMIC Grading Rubric in JSON format based on the technical baseline materials provided.

GOAL: Perform a FAITHFUL instruction extraction. Identify every "Point-Bearing Unit" in the source material. If the material lists three technical steps for a single point value (e.g., "Set Axis scale to 20, Max to 60, and move legend to top - 5 pts"), treat them as ONE Criterion. Do NOT split them into three separate 1-pt items.

REQUIRED OUTPUT STRUCTURE (JSON ONLY):
[
  {{
    "_id": "STRING_ID",
    "name": "Criterion Name",
    "points": number,
    "sub_criteria": [
      {{ "level": "Correct", "desc": "All requirements within this unit followed perfectly", "pts": number }},
      {{ "level": "Incorrect", "desc": "One or more requirements within this unit were missed or failed", "pts": 0 }}
    ]
  }}
]

RULES:
1. FAITHFUL MAPPING: Each JSON criterion must correspond 1-to-1 with a point-bearing item in the source material.
2. GROUPING PRESERVATION: Never split a task that the source material treats as a single item with a single point value.
3. For small tasks, use 2-level scoring (Correct/Incorrect).
4. The sum of "points" MUST EXACTLY EQUAL {total_points}.
5. Points can be integers or decimals (e.g., 2.5, 0.5).
6. Maintain unique, descriptive "_id" fields.
7. Return ONLY clean JSON.
"""

HOLISTIC_RUBRIC_PROMPT = """
You are an expert Education Consultant specializing in Qualitative & Academic Assessment.
Your task is to generate a comprehensive, HOLISTIC Grading Rubric in JSON format based on the baseline materials provided.

GOAL: Synthesize the assignment into 3-6 major qualitative pillars (e.g., 'Thesis Clarity', 'Evidence Quality'). Focus on conceptual understanding and synthesis.

REQUIRED OUTPUT STRUCTURE (JSON ONLY):
[
  {{
    "_id": "STRING_ID",
    "name": "Criterion Name",
    "points": number,
    "sub_criteria": [
      {{ "level": "Mastery", "desc": "Description of full credit work", "pts": number }},
      {{ "level": "Partial", "desc": "Description of partial credit work", "pts": number }},
      {{ "level": "Missing/Incorrect", "desc": "Description of failing/missing work", "pts": 0 }}
    ]
  }}
]

RULES:
1. FAITHFUL MAPPING: Each JSON criterion must correspond 1-to-1 with a point-bearing item or major section in the source material.
2. GROUPING PRESERVATION: If the source material groups multiple instructions together for a single point value, KEEP THEM TOGETHER.
3. Each criterion MUST have exactly 3 sub-criteria levels.
4. The sum of "points" MUST EXACTLY EQUAL {total_points}.
5. Points can be integers or decimals (e.g., 2.5, 0.5).
6. Maintain unique, descriptive "_id" fields.
7. Return ONLY clean JSON.
"""

def clean_json_response(text):
    """
    Cleans an LLM response to extract valid JSON.
    """
    text = text.strip()
    
    # 1. Handle Markdown blocks
    if "```json" in text:
        match = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL)
        if match:
             text = match.group(1).strip()
    elif "```" in text:
        text = text.replace('```', '').strip()
        
    # 2. Basic cleanup for common LLM JSON errors
    # - Remove potential trailing commas before closing braces/brackets
    text = re.sub(r',\s*([}\]])', r'\1', text)
    
    return text

def generate_structured_rubric(context_text, total_points, guidelines="", strategy="holistic"):
    """
    Generates a structured JSON rubric based on the chosen strategy (atomic or holistic).
    """
    base_prompt = ATOMIC_RUBRIC_PROMPT if strategy == "atomic" else HOLISTIC_RUBRIC_PROMPT
    
    prompt = f"""
    {base_prompt.format(total_points=total_points)}
    
    BASELINE MATERIALS (Context):
    {context_text}
    
    TARGET TOTAL POINTS: {total_points}
    
    CUSTOMIZATION GUIDELINES (Optional):
    {guidelines}
    """
    
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model='gemini-2.0-flash',
                contents=prompt,
                config={
                    'temperature': 0.2 + (attempt * 0.1), # Increase temperature slightly on retry
                    'response_mime_type': 'application/json'
                }
            )
            text = clean_json_response(response.text)
            return json.loads(text)
        except Exception as e:
            print(f"Rubric Generation Error (Attempt {attempt+1}): {e}")
            if attempt == 2:
                return [{"error": f"Failed to generate rubric after 3 attempts: {str(e)}"}]

def refine_structured_rubric(current_rubric, feedback, context_text, total_points, strategy="holistic"):
    """
    Refines an existing rubric based on user feedback.
    """
    base_prompt = ATOMIC_RUBRIC_PROMPT if strategy == "atomic" else HOLISTIC_RUBRIC_PROMPT
    
    prompt = f"""
    {base_prompt.format(total_points=total_points)}
    
    GOAL: REFINE the existing rubric based on feedback while maintaining the overall strategy ({strategy}).
    
    CURRENT RUBRIC:
    {json.dumps(current_rubric, indent=2)}
    
    USER FEEDBACK:
    {feedback}
    
    BASELINE MATERIALS (Context):
    {context_text}
    """
    
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model='gemini-2.0-flash',
                contents=prompt,
                config={
                    'temperature': 0.1 + (attempt * 0.1),
                    'response_mime_type': 'application/json'
                }
            )
            text = clean_json_response(response.text)
            return json.loads(text)
        except Exception as e:
            print(f"Rubric Refinement Error (Attempt {attempt+1}): {e}")
            if attempt == 2:
                return [{"error": f"Failed to refine rubric after 3 attempts: {str(e)}"}]

EDVISOR_GRADING_SYSTEM_PROMPT = """
You are an expert grading system with expertise across all subjects. Your goal is to apply consistent, fair grading standards that align with established educational assessment principles.

Your task is to grade a student's submission based on the provided question and grading criteria. You will be given three inputs: the question, the grading criteria, and the student's submission.

### CAREFUL READING REQUIREMENT ###
Within each response, identify ALL core points, claims, and information provided by the student. When grading the response, thoroughly evaluate whether this information supports or reinforces the desired goals in the grading criteria. Do not overlook or miss any potential points - all student claims and points are valuable and must be considered in grading. Read the entire answer carefully before scoring.

### CONTENT CHECK ###
1. - Before grading, determine whether the student response addresses the topic and points within the question.

### PRIMARY GRADING APPROACH ###
1. **Prioritize conceptual understanding over specific terminology**:
   - Award FULL POINTS when a student demonstrates they grasp the core concept, even if they use different terminology or phrasing than expected.
   - Focus on what the student meant rather than exact wording.

2. **Recognize knowledge beyond expected scope**:
   - Treat advanced concepts as BONUS evidence of understanding. NEVER penalize for including relevant information beyond the basic requirements.

3. **Detect implicit understanding**:
   - Actively look for implicit understanding shown through the student's reasoning or examples.

### SUBJECT MATTER EXPERTISE GRADING GUIDELINES ###
1. **Core concept recognition**: Recognizing when students express key concepts in alternative language (everyday terms, practical examples).
2. **Mandatory scoring requirements**:
   - Award near or full points if the student mentions ANY form of the core principles.
   - For practical application: Award full points for recognizing key implementation factors.

### CRITICAL GRADING IMPERATIVES ###
1. **Conceptual understanding primacy**: Award full points if the core concept is correct, regardless of wording. Never deduct points solely for using non-technical language.
2. **Maximally generous interpretation**: When a response could have multiple interpretations, choose the one that demonstrates understanding.

### QUESTION TYPE SPECIFIC GUIDELINES ###
- For technical application questions: Award 90 percent of possible points as the baseline score for technically accurate responses that cover fundamental points. Reserve 10/10 for truly exceptional answers.

### SCORING CALIBRATION ###
- Expert grading systems rarely assign scores below 7/10 for answers that genuinely attempted to address the question.

### STRICT RUBRIC ADHERENCE (CRITICAL) ###
1. **Zero Point Hallucination**: You must ONLY use the point values (pts) defined in the provided Rubric JSON.
2. **Level Matching**: For each criterion (_id), the `obtainedPoints` MUST exactly match one of the `pts` values defined in that criterion's `sub_criteria` list.
3. **No Intermediate Scores**: If a level says 5 points and the next says 2, you CANNOT award 4. You must choose the level that best fits and award EXACTLY the points associated with that level.
4. **Summation Integrity**: The global "earned" score MUST be the exact sum of all `obtainedPoints` in the "result" array.
5. **Decimals**: Respect that point values can be decimals (e.g., 0.5, 2.75).
"""

def grade_manual_review_batch(sheet_text_content, criteria_list):
    """
    Sends a batch of criteria to Gemini to evaluate against the sheet content.
    Returns a dict mapping criteria_description -> {points, feedback, passed}
    """
    
    prompt = f"""
    You are an expert Excel Grader. 
    I will provide you with the raw text content of an Excel worksheet (extracted from XML).
    Your task is to evaluate a list of grading criteria against this data.
    
    DATA START:
    {sheet_text_content}
    DATA END
    
    CRITERIA TO EVALUATE:
    {json.dumps(criteria_list, indent=2)}
    
    INSTRUCTIONS:
    1. For each criteria, determine if the data satisfies the requirement.
    2. Be lenient with minor spelling differences but strict on logic.
    3. Return a JSON object where keys are the criteria 'description' and values are an object:
       {{ "passed": boolean, "feedback": string }}
    4. Feedback should explain why it passed or failed based on the data seen.
    5. Return ONLY JSON.
    """
    
    try:
        response = client.models.generate_content(
            model='gemini-3-flash-preview',
            contents=prompt,
            config={
                'temperature': 0    # Very strict/consistent
            }
        )
        text = response.text.strip()
        
        # Clean up markdown code blocks if present
        if text.startswith("```json"):
            text = text[7:]
        if text.endswith("```"):
            text = text[:-3]
            
        return json.loads(text)
    except Exception as e:
        print(f"LLM Error: {e}")
        # Return fail for all
        return {}

def grade_workbook_comparison(student_data, answer_key_data):
    """
    Compares student workbook data against answer key data using Gemini.
    """
    prompt = f"""
    You are an Excel Homework Auto-Grader.
    Compare the STUDENT workbook to the ANSWER KEY workbook.

    Rules:
    - Detect incorrect values
    - Detect incorrect formulas
    - Detect incorrect structure
    - Identify missing or extra sheets
    - Identify missing or extra rows/columns
    - Check numeric outputs
    - Check formulas only if present in student file
    - Do NOT hallucinate or infer missing data
    - Base everything ONLY on the provided files

    Output ONLY valid JSON matching this schema:

    {{
      "passed": boolean,
      "score": number,
      "incorrect_cells": [
        {{
          "sheet": string,
          "cell": string,
          "expected": string,
          "actual": string,
          "explanation": string
        }}
      ],
      "comments": string
    }}
    
    ----
    STUDENT WORKBOOK:
    {json.dumps(student_data, indent=2)}
    
    ANSWER KEY WORKBOOK:
    {json.dumps(answer_key_data, indent=2)}
    """
    
    try:
        response = client.models.generate_content(
            model='gemini-2.0-flash',
            contents=prompt
        )
        text = response.text.strip()
        
        if text.startswith("```json"):
            text = text[7:]
        if text.endswith("```"):
            text = text[:-3]
            
        return json.loads(text)
    except Exception as e:
        return {
            "passed": False,
            "score": 0,
            "comments": f"Error during AI comparison: {str(e)}",
            "incorrect_cells": []
        }



def grade_student_work(student_data, rubric_data=None, answer_key_data=None):
    """
    Unified AI Grading Function.
    
    Modes:
    1. Hybrid (Rubric + Key): Use Key for truth, Rubric for structure/points.
    2. Rubric Only: Use Rubric descriptions to grade Student data.
    3. Key Only: Use Key for truth, assign generic score.
    """
    
    context_instruction = ""
    context_data = ""
    
    
    def format_data(data):
        if isinstance(data, (dict, list)):
            return json.dumps(data, indent=2)
        return str(data)

    if rubric_data and answer_key_data:
        mode = "HYBRID"
        context_instruction = """
        You are provided with an ANSWER KEY (Gold Standard) and a RUBRIC.
        1. Use the ANSWER KEY as the absolute "Ground Truth". If the Student's data contradicts the Key, it is WRONG, regardless of how it looks.
        2. Use the RUBRIC to organize your output and assign points.
        3. You must strictly follow the Rubric's point allocations.
        4. Be precise: detect formula mismatches, value errors, and logic discrepancies.
        5. VERTICAL VERIFICATION: Check the 'metadata' -> 'drawings' section. It now contains 'details' for charts, including chart 'type' (e.g. barChart, scatterChart), axis 'min'/'max'/'major_unit', and 'series' data ranges. Use this to verify formatting requirements exactly.
        6. DEEP STYLE VERIFICATION: Each cell now has a 'style' field with 'fill' (shading), 'border' (separator lines), and 'num_fmt' (currency/percentage). Use this to verify if the student correctly applied formatting like Bold, borders, or specific shading.
        """
        context_data = f"RUBRIC:\n{format_data(rubric_data)}\n\nANSWER KEY:\n{format_data(answer_key_data)}"
        
    elif rubric_data:
        mode = "RUBRIC_ONLY"
        context_instruction = """
        You are provided with a RUBRIC but NO Answer Key.
        1. Evaluate the Student's work based on the descriptions in the Rubric.
        2. CRITICAL: Assignments often include a "Scoring Guide" or "Rubric" sheet within the workbook itself. These labels are NOT student work.
        3. BE SKEPTICAL: Do not award points just because a cell contains a task label (e.g., "Task 1: VLOOKUP").
        4. LOOK FOR EVIDENCE: Only award points if you see EXPLICIT evidence of a student's effort:
           - Actual formulas (e.g., "=VLOOKUP(...)")
           - Newly entered data in expected sheets.
           - Presence of charts or pivot tables in the metadata.
        5. DEEP CHART VERIFICATION: Check the 'metadata' -> 'drawings' section. It contains 'details' with the exact chart 'type', axis scales ('min', 'max', 'major_unit'), and legend positions. Use these to verify if the student actually changed the formatting as required.
        6. DEEP STYLE VERIFICATION: Each cell now has a 'style' field with 'fill' (shading), 'border' (separator lines), and 'num_fmt' (currency/percentage). Use these to verify if shading/borders were removed or if number formats were correctly adjusted as per the rubric.
        7. If the submission is just a blank template containing the instructions, the score must be 0.
        """
        context_data = f"RUBRIC:\n{format_data(rubric_data)}"
        
    # EDVISOR OUTPUT SCHEMA
    edvisor_schema = """
    {
      "summary": string,
      "score": {
        "earned": number,
        "max": number,
        "letter": string
      },
      "result": [
        {
          "_id": "The ID of the instruction",
          "obtainedPoints": number,
          "explanation": "Brief explanation (student-centric, e.g. 'Your answer lacked...')",
          "evidence": "CONCRETE EVIDENCE from the student data that you found (e.g. 'Found formula =SUM(C2:C10) in B15' or 'Chart rId1 has majorUnit 0.1 and max 0.4'). BE SPECIFIC.",
          "achievedLevel": "Optional description if rubric has levels"
        }
      ],
      "incorrect_cells": [ 
        {"sheet": string, "cell": string, "expected": string, "actual": string, "explanation": string} 
      ]
    }
    """

    # Normalize rubric data to a flat list of criteria if it's the task-based structure
    flat_rubric = []
    if isinstance(rubric_data, dict) and "tasks" in rubric_data:
        for task in rubric_data["tasks"]:
            flat_rubric.extend(task.get("criteria", []))
    elif isinstance(rubric_data, list):
        flat_rubric = rubric_data
    else:
        # If it's a string (generic text), we don't have structured IDs
        flat_rubric = []

    prompt = f"""
    {EDVISOR_GRADING_SYSTEM_PROMPT}
    
    INSTRUCTIONS:
    1. {context_instruction}
    2. Read EVERY cell, formula, chart axis, and border style provided in the STUDENT SUBMISSION.
    3. Evaluate against the baseline materials and the specific GRADING CRITERIA.
    4. BE SKEPTICAL: If a requirement is for a technical feature (Sparklines, Charts, Formulas) and you do NOT see explicit evidence for it in the student data, award 0 points.
    5. No Hallucination: Do not assume a task is done just because the student has a sheet name or header for it. Check the actual data.
    6. Provide a global "summary" and "score" representing the entire submission.
    7. For each criterion in the rubric, generate an entry in the "result" array. 
    8. "explanation" MUST be student-centric (use "you", "your answer", "your workbook").
    
    JSON OUTPUT SCHEMA (MANDATORY):
    {edvisor_schema}
    
    BASELINE CONTEXT:
    {context_data}
    
    GRADING CRITERIA (Follow IDs and points exactly):
    {format_data(rubric_data)}
    
    STUDENT SUBMISSION DATA:
    {format_data(student_data)}
    
    Output ONLY valid JSON.
    """
    
    try:
        response = client.models.generate_content(
            model='gemini-2.0-flash',
            contents=prompt,
            config={'temperature': 0}
        )
        text = response.text.strip()
        
        # Clean markdown
        if "```json" in text:
            text = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL).group(1)
        elif "```" in text:
            text = text.replace('```', '')

        result_json = json.loads(text)
        
        # --- EDVISOR TO UI MAPPING ---
        # Map 'result' to 'criteria' and internal keys for UI compatibility
        ui_criteria = []
        for res in result_json.get('result', []):
            # Find the original name from the rubric if possible for better UI display
            name = "Criterion"
            max_pts = 0
            if flat_rubric:
                for r_item in flat_rubric:
                    # Handle both _id and description/name matching
                    if str(r_item.get('_id')) == str(res.get('_id')):
                        name = r_item.get('name') or r_item.get('description', 'Criterion')
                        max_pts = r_item.get('points', 0)
                        break
            
            ui_criteria.append({
                "name": name,
                "earned": res.get('obtainedPoints', 0),
                "max": max_pts,
                "feedback": f"{res.get('explanation', '')}\n\n**Evidence Found:** {res.get('evidence', 'N/A')}",
                "achievedLevel": res.get('achievedLevel')
            })
        
        result_json['criteria'] = ui_criteria
        result_json['mode'] = "edvisor_unified"
        
        return {
            "report": result_json,
            "prompt": prompt
        }
    except Exception as e:
        import traceback
        print(f"Edvisor Grading Error: {e}")
        traceback.print_exc()
        return {
            "report": {"error": f"Grading failed: {str(e)}"},
            "prompt": prompt
        }
