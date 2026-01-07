from utils.xml_helper import get_sheet_map, get_shared_strings, parse_sheet_full
import re

def extract_rubric_from_sheet(unzip_dir):
    """
    Attempts to find a 'Scoring Guide' or 'Rubric' sheet and parse it.
    Returns a dict with 'source': 'embedded', 'tasks': [...] or None.
    """
    sheet_map = get_sheet_map(unzip_dir)
    shared_strings = get_shared_strings(unzip_dir)
    
    target_sheet_key = None
    for k in sheet_map.keys():
        if "scoring" in k.lower() or "rubric" in k.lower():
            target_sheet_key = k
            break
            
    if not target_sheet_key:
        return None
        
    xml_path = sheet_map[target_sheet_key]
    data, _ = parse_sheet_full(xml_path, shared_strings)
    
    # Sort items by row then col
    def parse_coord(coord):
        match = re.match(r"([A-Z]+)(\d+)", coord)
        if match:
            return match.group(1), int(match.group(2))
        return "A", 0
        
    sorted_cells = sorted(data.items(), key=lambda x: parse_coord(x[0])[1])
    
    tasks = []
    current_task = None
    
    # Heuristic: 
    # Col C = Description (Task Name if points empty, Criteria text if points exist)
    # Col D = Points
    
    # We group by rows.
    rows = {}
    for coord, val in sorted_cells:
        col, row = parse_coord(coord)
        if row not in rows:
            rows[row] = {}
        rows[row][col] = val['value']
        
    sorted_row_nums = sorted(rows.keys())
    
    current_task = {
        "name": "General",
        "sheet": "Introduction", # Fallback
        "points": 0,
        "criteria": []
    }
    tasks.append(current_task)
    
    for r in sorted_row_nums:
        row_data = rows[r]
        desc = row_data.get('C', '').strip()
        points_str = row_data.get('D', '').strip()
        
        if not desc:
            continue
            
        # Skip summary lines
        if "possible/deducted" in desc.lower():
            continue
            
        # Try to parse points
        points = 0
        try:
            points = float(points_str)
        except ValueError:
            points = 0
            
        if points > 0:
            # It's a criteria item
            current_task['criteria'].append({
                "type": "manual_review", # We can't know the logic automatically
                "description": desc,
                "expected": "See description",
                "cell": "N/A",
                "points": points,
                "feedback_on_fail": f"Check: {desc}"
            })
            current_task['points'] += points
        elif len(desc) > 5 and not desc.lower().startswith('note'):
            # Likely a Task Header if no points, e.g. "T 1 - VLOOKUP"
            # Try to match it to a known sheet
            matched_sheet = None
            for sname in sheet_map.keys():
                # Fuzzy match: if header contains sheet name or vice versa (simplified)
                # Remove special chars for comparison
                clean_desc = re.sub(r'[^a-zA-Z0-9]', '', desc).lower()
                clean_sname = re.sub(r'[^a-zA-Z0-9]', '', sname).lower()
                
                if clean_desc in clean_sname or clean_sname in clean_desc:
                    matched_sheet = sname
                    break
            
            # Start new task if matched sheet OR if it looks like a task header
            is_task_header = re.search(r"^(T|Task)\s*\d", desc, re.IGNORECASE)
            
            if matched_sheet or is_task_header:
                current_task = {
                    "name": desc,
                    "sheet": matched_sheet if matched_sheet else current_task['sheet'], # Inherit sheet if unknown
                    "points": 0,
                    "criteria": []
                }
                tasks.append(current_task)
            else:
                 # Just update name of current generic task if it's empty
                 if current_task['name'] == "General" and not current_task['criteria']:
                     current_task['name'] = desc
    
    # Filter empty tasks
    final_tasks = [t for t in tasks if t['criteria']]
    
    # Add unique IDs if missing
    for i, task in enumerate(final_tasks):
        for j, crit in enumerate(task['criteria']):
            if '_id' not in crit:
                crit['_id'] = f"T{i+1}_crit_{j+1}"
            if 'name' not in crit:
                # Use description as name if too short, or truncate
                crit['name'] = crit['description'][:50] + ("..." if len(crit['description']) > 50 else "")

    if not final_tasks:
        return None
        
    return {"tasks": final_tasks, "source": "embedded_scanned"}
