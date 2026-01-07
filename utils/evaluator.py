from utils.llm_helper import grade_manual_review_batch
import re
import json

def check_value_match(cell_data, expected_value, tolerance=None):
    """
    Checks if cell value matches expected.
    If numbers, can use tolerance.
    """
    actual = cell_data.get('value', '')
    
    # Try numeric comparison if both look like numbers
    try:
        f_actual = float(actual)
        f_expected = float(expected_value)
        if tolerance is not None:
            return abs(f_actual - f_expected) <= tolerance
        return f_actual == f_expected
    except ValueError:
        # String comparison
        return str(actual).strip().lower() == str(expected_value).strip().lower()

def check_formula_match(cell_data, expected_substring, is_regex=False):
    """
    Checks if formula contains substring or matches regex.
    """
    actual_formula = cell_data.get('formula', '')
    if not actual_formula:
        return False
        
    if is_regex:
        return re.search(expected_substring, actual_formula, re.IGNORECASE) is not None
    else:
        return expected_substring.lower() in actual_formula.lower()

def evaluate_criteria(sheet_data, criteria):
    """
    Evaluates a single criteria object against the sheet data.
    Returns (points_earned, feedback_list)
    """
    cell_ref = criteria.get('cell')
    ctype = criteria.get('type')
    points = criteria.get('points', 0)
    
    # Bypass cell check for manual review or other types that don't need a specific cell
    if ctype == 'manual_review':
        # This is now handled in batch by the task evaluator, but if called individually,
        # we return 0/Pending if logic not here.
        # Ideally, we shouldn't call this individually for manual_review if we want batching.
        return 0, ["Processed in batch"]
        
    if cell_ref not in sheet_data:
        return 0, [f"Cell {cell_ref} not found or empty."]
        
    cell = sheet_data.get(cell_ref, {})
    passed = False
    
    if ctype == 'value_match':
        passed = check_value_match(cell, criteria.get('expected'), criteria.get('tolerance'))
    elif ctype == 'formula_match':
        passed = check_formula_match(cell, criteria.get('expected'), criteria.get('regex', False))
    elif ctype == 'exists':
        passed = True # Cell exists check passed effectively by 'if cell_ref not in sheet_data' above
    else:
        return 0, [f"Unknown criteria type: {ctype}"]
        
    if passed:
        return points, []
    else:
        feedback = criteria.get('feedback_on_fail', f"Check {ctype} in {cell_ref} failed.")
        return 0, [feedback]

def evaluate_task(task, sheet_data, sheet_metadata=None):
    """
    Evaluates a full task (list of criteria).
    """
    task_name = task.get('name', 'Unknown Task')
    max_points = task.get('points', 0)
    criteria_list = task.get('criteria', [])
    
    earned_points = 0
    criteria_results = []
    all_feedback = []
    
    # Separate logic for batch LLM processing
    manual_criteria = [c for c in criteria_list if c.get('type') == 'manual_review']
    automated_criteria = [c for c in criteria_list if c.get('type') != 'manual_review']
    
    # 1. Process Automated
    for crit in automated_criteria:
        p, fb = evaluate_criteria(sheet_data, crit)
        earned_points += p
        all_feedback.extend(fb)
        
        desc = crit.get('description')
        if not desc:
            desc = f"Check {crit.get('type')} at {crit.get('cell')}"
            
        criteria_results.append({
            "description": desc,
            "points_earned": p,
            "max_points": crit.get('points', 0),
            "status": "Passed" if p == crit.get('points', 0) else "Failed",
            "feedback": "; ".join(fb) if fb else "Correct"
        })
        
    # 2. Process Manual (Batch)
    if manual_criteria:
        # Convert sheet data to text representation
        # We'll dump non-empty cells relevant for context
        # Sorting helps LLM read it like a grid
        def parse_coord(coord):
            match = re.match(r"([A-Z]+)(\d+)", coord)
            if match:
                return match.group(1), int(match.group(2))
            return "A", 0
            
        sorted_cells = sorted(sheet_data.items(), key=lambda x: parse_coord(x[0])[1])
        sheet_text = []
        for coord, meta in sorted_cells:
            val = meta.get('value', '')
            formula = meta.get('formula')
            if val or formula:
                content = f"{coord}: {val}"
                if formula:
                    content += f" (Formula: {formula})"
                sheet_text.append(content)
        
        full_text = "\n".join(sheet_text)
        
        # Add metadata context if provided
        if sheet_metadata:
             full_text += "\n\n[METADATA START]\n"
             full_text += json.dumps(sheet_metadata, indent=2)
             full_text += "\n[METADATA END]\n"
        
        # Call LLM
        llm_results = grade_manual_review_batch(full_text, manual_criteria)
        
        for crit in manual_criteria:
            desc = crit.get('description')
            points = crit.get('points', 0)
            
            res = llm_results.get(desc, {"passed": False, "feedback": "LLM Verification Failed"})
            
            p_earned = 0
            if res.get('passed'):
                p_earned = points
                earned_points += points
            else:
                fb = res.get('feedback', 'Criteria not met.')
                all_feedback.append(f"[AI REVIEW] {desc}: {fb}")
                
            criteria_results.append({
                "description": desc,
                "points_earned": p_earned,
                "max_points": points,
                "status": "Passed" if res.get('passed') else "Failed",
                "feedback": res.get('feedback', 'Criteria not met.')
            })

    status = "Full Credit"
    if earned_points == 0:
        status = "No Credit"
    elif earned_points < max_points - 0.01: # Float safety
        status = "Partial"
        
    return {
        "task_name": task_name,
        "points_earned": round(earned_points, 2),
        "max_points": max_points,
        "status": status,
        "feedback": "; ".join(all_feedback) if all_feedback else "Correct",
        "criteria_results": criteria_results
    }
