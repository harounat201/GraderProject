def test_interactive_logic():
    print("--- Testing Interactive Rubric Logic ---")
    
    # 1. Mock Step 2: Manual Entry
    manual = [{"name": "Manual Task", "points": 10, "sub_criteria": []}]
    total_goal = 50
    manual_sum = sum(c['points'] for c in manual)
    remaining = total_goal - manual_sum
    print(f"Goal: {total_goal}, Manual: {manual_sum}, AI target: {remaining}")
    
    # 2. Mock Step 3: Merging
    ai_generated = [{"name": "AI Task", "points": remaining, "sub_criteria": []}]
    merged = manual + ai_generated
    print(f"Merged count: {len(merged)}")
    current_sum = sum(c['points'] for c in merged)
    print(f"Final Sum: {current_sum} (Matches Goal: {current_sum == total_goal})")

    # 3. Mock Step 3: Editing
    merged[0]['points'] = 20 # User edits manual task points
    new_sum = sum(c['points'] for c in merged)
    print(f"After Edit Sum: {new_sum} (Difference: {new_sum - total_goal})")

if __name__ == "__main__":
    test_interactive_logic()
