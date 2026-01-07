from utils.xml_helper import get_sheet_map, get_shared_strings, parse_sheet_data
import zipfile
import tempfile
import shutil
import os

f = "DataManagement.xlsm"
t = tempfile.mkdtemp()
with zipfile.ZipFile(f, 'r') as z:
    z.extractall(t)

s_map = get_sheet_map(t)
ss = get_shared_strings(t)

# Find sheet with "score" or "rubric" in name
target_sheet = None
for k in s_map.keys():
    if "scoring" in k.lower() or "rubric" in k.lower():
        target_sheet = k
        break

if target_sheet:
    print(f"Found Rubric Sheet: {target_sheet}")
    xml = s_map[target_sheet]
    data = parse_sheet_data(xml, ss)
    # Dump first 50 items to see structure
    # Sort keys to make it readable row-by-row
    sorted_items = sorted(data.items(), key=lambda x: (int(''.join(filter(str.isdigit, x[0])) or 0), x[0]))
    for k, v in sorted_items[:50]:
        print(f"{k}: {v['value']}")
else:
    print("No rubric sheet found.")

shutil.rmtree(t)
