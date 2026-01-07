import streamlit as st
import json
import os
import tempfile
from grader import grade_submission, prepare_grading_context
from utils.llm_helper import generate_structured_rubric, refine_structured_rubric

st.set_page_config(page_title="Assignment Autograder Workflow Prototype", page_icon="📝", layout="wide")

# --- SESSION STATE INITIALIZATION ---
if 'grading_step' not in st.session_state:
    st.session_state.grading_step = 1
if 'context_files' not in st.session_state:
    st.session_state.context_files = []
if 'generated_rubric' not in st.session_state:
    st.session_state.generated_rubric = []
if 'total_points' not in st.session_state:
    st.session_state.total_points = 100.0
if 'custom_guidelines' not in st.session_state:
    st.session_state.custom_guidelines = ""
if 'manual_criteria' not in st.session_state:
    st.session_state.manual_criteria = []
if 'manual_mode' not in st.session_state:
    st.session_state.manual_mode = False
if 'grading_strategy' not in st.session_state:
    st.session_state.grading_strategy = "holistic"

st.title("📝 Assignment Autograder PROMPT Playground")

# --- STEP 1: GRADING CONTEXT ---
if st.session_state.grading_step == 1:
    st.header("Step 1: Setup Grading Context")
    st.write("Upload the Rubrics, Answer Keys, or Prompts you want the AI to use as a baseline.")
    
    context_uploads = st.file_uploader(
        "Upload Baseline Materials", 
        type=["json", "xlsx", "xlsm", "docx", "txt"], 
        accept_multiple_files=True,
        key="context_uploader"
    )
    
    st.session_state.total_points = st.number_input("Total Assignment Points", min_value=0.5, value=float(st.session_state.total_points), step=0.5)

    c1, c2 = st.columns(2)
    if c1.button("Confirm Uploads"):
        if not context_uploads:
            st.warning("Please upload at least one context file.")
        else:
            st.session_state.context_files = context_uploads
            # Strategy Detection
            has_excel = any(f.name.lower().endswith(('.xlsx', '.xlsm')) for f in context_uploads)
            st.session_state.grading_strategy = "atomic" if has_excel else "holistic"
            
            st.session_state.manual_mode = False
            st.session_state.grading_step = 2
            st.rerun()

    if c2.button("Skip AI & Build Manually"):
        st.session_state.context_files = context_uploads if context_uploads else []
        st.session_state.manual_mode = True
        st.session_state.grading_step = 3
        st.rerun()

# --- STEP 2: CUSTOMIZATION & MANUAL ENTRY ---
if st.session_state.grading_step == 2:
    st.header("Step 2: Customization & Manual Entry")
    
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.write("### Baseline Files:")
        for f in st.session_state.context_files:
            st.write(f"- {f.name}")
        st.write(f"**Target Total Points**: {st.session_state.total_points}")
        
        st.divider()
        st.write("### ➕ Manual Criteria (Optional)")
        st.write("Add specific criteria you want to include before AI generation.")
        
        with st.form("manual_crit_form"):
            new_crit_name = st.text_input("Criterion Name", placeholder="e.g., Specific Formatting Task")
            new_crit_pts = st.number_input("Points", min_value=0.5, value=10.0, step=0.5)
            if st.form_submit_button("Add Manual Criterion"):
                if new_crit_name:
                    st.session_state.manual_criteria.append({
                        "name": new_crit_name,
                        "points": new_crit_pts,
                        "sub_criteria": [
                            {"level": "Mastery", "desc": "Full requirements met.", "pts": new_crit_pts},
                            {"level": "Partial", "desc": "Partial requirements met.", "pts": int(new_crit_pts/2)},
                            {"level": "Missing", "desc": "Not attempted.", "pts": 0}
                        ]
                    })
                    st.rerun()

    with col_b:
        st.write("### Current List:")
        if not st.session_state.manual_criteria:
            st.info("No manual criteria added yet.")
        else:
            for i, mc in enumerate(st.session_state.manual_criteria):
                c1, c2 = st.columns([3, 1])
                c1.write(f"{mc['name']} ({mc['points']} pts)")
                if c2.button("🗑️", key=f"del_{i}"):
                    st.session_state.manual_criteria.pop(i)
                    st.rerun()

    st.divider()
    @st.dialog("Generate Final Rubric")
    def customization_dialog():
        st.write("Enter any specific instructions or focus areas for the AI (optional).")
        guidelines = st.text_area("Customization Guidelines", value=st.session_state.custom_guidelines, placeholder="e.g., Focus heavily on APA citations and thesis clarity.")
        if st.button("Final Generate & Merge"):
            st.session_state.custom_guidelines = guidelines
            st.session_state.grading_step = 3
            st.rerun()

    c1, c2, c3 = st.columns([1, 1, 3])
    if c1.button("← Back"):
        st.session_state.grading_step = 1
        st.rerun()
    if c2.button("Generate Rubric", type="primary"):
        customization_dialog()

# --- STEP 3: CRITERIA DISPLAY ---
if st.session_state.grading_step == 3:
    st.header("Step 3: Review & Edit Grading Criteria")
    
    # Generate from context if empty and not in manual mode
    if not st.session_state.generated_rubric:
        if st.session_state.manual_mode:
            # Initialize with one blank criterion so the UI isn't empty
            st.session_state.generated_rubric = [{
                "_id": "manual_1",
                "name": "New Criterion",
                "points": 10,
                "sub_criteria": [
                    {"level": "Mastery", "desc": "Full credit.", "pts": 10},
                    {"level": "Partial", "desc": "Partial credit.", "pts": 5},
                    {"level": "Missing", "desc": "No credit.", "pts": 0}
                ]
            }]
        else:
            with st.spinner("Analyzing context and generating structured rubric..."):
                # Calculate remaining points for AI to distribute
                manual_sum = sum(c['points'] for c in st.session_state.manual_criteria)
                remaining_pts = max(0, st.session_state.total_points - manual_sum)
                
                context_text = prepare_grading_context(st.session_state.context_files)
                # Instruct AI to only generate for the remaining points
                ai_rubric = generate_structured_rubric(
                    context_text, 
                    remaining_pts, 
                    st.session_state.custom_guidelines,
                    strategy=st.session_state.grading_strategy
                )
                
                # Merge: Manual first, then AI
                st.session_state.generated_rubric = st.session_state.manual_criteria + ai_rubric

    if st.session_state.generated_rubric and "error" in st.session_state.generated_rubric[0]:
        st.error(st.session_state.generated_rubric[0]["error"])
        if st.button("Retry"):
            st.session_state.generated_rubric = []
            st.rerun()
    else:
        # POINT TRACKER
        current_total = round(sum(c.get('points', 0) for c in st.session_state.generated_rubric), 2)
        diff = round(current_total - st.session_state.total_points, 2)
        
        col1, col2, col3 = st.columns([1,1,1])
        with col1:
            new_goal = st.number_input("Goal Points", min_value=0.5, value=float(st.session_state.total_points), step=0.5, key="goal_pts_step3")
            if new_goal != st.session_state.total_points:
                st.session_state.total_points = new_goal
                st.rerun()
        
        col2.metric("Current Sum", current_total, delta=-diff if diff != 0 else 0, delta_color="inverse" if diff > 0 else "normal")
        
        if diff != 0:
            st.warning(f"⚠️ Points mismatch! Your criteria add up to {current_total}, but the goal is {st.session_state.total_points}.")

        st.divider()

        # EDITABLE LIST
        for i, crit in enumerate(st.session_state.generated_rubric):
            with st.expander(f"📝 {crit.get('name', 'Criterion')} ({crit.get('points', 0)} pts)", expanded=(i == len(st.session_state.generated_rubric)-1)):
                # Main Criterion Edit
                c_name = st.text_input("Name", value=crit.get('name', ''), key=f"name_{i}")
                c_pts = st.number_input("Points", value=float(crit.get('points', 0)), step=0.5, key=f"pts_{i}")
                
                # Update state immediately
                st.session_state.generated_rubric[i]['name'] = c_name
                st.session_state.generated_rubric[i]['points'] = c_pts
                
                st.write("---")
                st.write("**Sub-Criteria (Scoring Levels):**")
                
                # Sub-criteria edit
                for j, sub in enumerate(crit.get('sub_criteria', [])):
                    sc_col1, sc_col2, sc_col3 = st.columns([2, 5, 2])
                    s_level = sc_col1.text_input("Level", value=sub.get('level', ''), key=f"lvl_{i}_{j}")
                    s_desc = sc_col2.text_input("Description", value=sub.get('desc', ''), key=f"desc_{i}_{j}")
                    s_pts = sc_col3.number_input("Pts", value=float(sub.get('pts', 0)), step=0.5, key=f"lvl_pts_{i}_{j}")
                    
                    # Update state
                    st.session_state.generated_rubric[i]['sub_criteria'][j] = {
                        "level": s_level, "desc": s_desc, "pts": s_pts
                    }
                
                if st.button("Delete Criterion", key=f"del_crit_{i}"):
                    st.session_state.generated_rubric.pop(i)
                    st.rerun()

        st.divider()
        st.write("### ✨ AI Refinement")
        st.write(f"Current Strategy: **{st.session_state.grading_strategy.capitalize()}**")
        st.session_state.grading_strategy = st.radio("Switch Strategy", ["holistic", "atomic"], index=0 if st.session_state.grading_strategy == "holistic" else 1, horizontal=True)
        
        st.write("Use natural language to adjust the entire rubric (e.g., 'Make Task 1 worth more points' or 'Add a focus on APA style').")
        refinement_feedback = st.text_area("Refinement Instructions", placeholder="Enter feedback here...", key="refinement_input")
        if st.button("Refine Rubric with AI"):
            if refinement_feedback:
                with st.spinner("Refining rubric based on your feedback..."):
                    context_text = prepare_grading_context(st.session_state.context_files)
                    refined = refine_structured_rubric(
                        st.session_state.generated_rubric,
                        refinement_feedback,
                        context_text,
                        st.session_state.total_points,
                        strategy=st.session_state.grading_strategy
                    )
                    if refined and "error" not in refined[0]:
                        st.session_state.generated_rubric = refined
                        st.success("Rubric refined successfully!")
                        st.rerun()
                    elif refined:
                        st.error(refined[0].get("error", "Refinement failed."))

        if st.button("➕ Add New Criterion"):
            st.session_state.generated_rubric.append({
                "name": "New Criterion",
                "points": 10.0,
                "sub_criteria": [
                    {"level": "Mastery", "desc": "Full credit.", "pts": 10.0},
                    {"level": "Partial", "desc": "Partial credit.", "pts": 5.0},
                    {"level": "Missing", "desc": "No credit.", "pts": 0.0}
                ]
            })
            st.rerun()
        
        c_back, c_next = st.columns([1, 4])
        if c_back.button("← Back / Regenerate"):
            st.session_state.generated_rubric = []
            st.session_state.grading_step = 2
            st.rerun()
        if c_next.button("Proceed to Grade Submission", type="primary"):
            if diff != 0:
                st.error("Please fix point totals before proceeding.")
            else:
                st.session_state.grading_step = 4
                st.rerun()

# --- STEP 4: FINAL GRADING ---
if st.session_state.grading_step == 4:
    st.header("Step 4: Grade Submission")
    
    uploaded_file = st.file_uploader("Upload Student Submission", type=["xlsx", "xlsm", "docx", "txt"])
    
    if st.button("Grade Submission", type="primary"):
        if not uploaded_file:
            st.error("Please upload a submission file.")
        else:
            with st.spinner("Grading..."):
                # Prepare inputs
                with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_path = tmp_file.name
                
                try:
                    # Execute grade using the generated rubric
                    full_results = grade_submission(
                        tmp_path, 
                        rubric_data=st.session_state.generated_rubric
                    )
                    
                    results = full_results.get("report", {})
                    raw_prompt = full_results.get("prompt", "")
                    
                    if "error" in results:
                        st.error(f"Error: {results['error']}")
                    else:
                        st.success("Grading Complete!")
                        
                        tab1, tab2, tab3 = st.tabs(["📝 Grading Report", "📄 Raw JSON", "🔍 AI Prompt"])
                        
                        with tab1:
                            score_data = results.get('score', {})
                            earned = score_data.get('earned', 0)
                            max_pts = score_data.get('max', st.session_state.total_points)
                            letter = score_data.get('letter', '')

                            col1, col2 = st.columns([1, 4])
                            with col1:
                                st.metric("Final Score", f"{earned} / {max_pts}", delta=letter)
                            with col2:
                                st.write("### AI Summary")
                                st.info(results.get('summary', 'No summary provided.'))
                            
                            st.divider()
                            st.subheader("Criteria Breakdown")
                            for crit in results.get('criteria', []):
                                c_name = crit.get('name', 'Criterion')
                                c_earn = crit.get('earned', 0)
                                c_max = crit.get('max', 0)
                                c_fb = crit.get('feedback', '')
                                icon = "✅" if c_earn == c_max else "⚠️" if c_earn > 0 else "❌"
                                with st.expander(f"{icon} {c_name} ({c_earn}/{c_max})"):
                                    st.write(f"**Feedback**: {c_fb}")
                        
                        with tab2:
                            st.json(results)
                        with tab3:
                            st.code(raw_prompt)
                            
                finally:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)

    if st.button("← Back to Criteria"):
        st.session_state.grading_step = 3
        st.rerun()
