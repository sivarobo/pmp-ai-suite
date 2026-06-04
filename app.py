import streamlit as st
import pandas as pd
from google import genai
from google.genai import types
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import re
import time

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

# ==========================================
# 1. API Configuration & Secrets Lock
# ==========================================
YOUR_API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=YOUR_API_KEY)

# ==========================================
# 2. Database Loading
# ==========================================
@st.cache_data
def load_data():
    try:
        return pd.read_csv('lesson_master_v1_5.csv')
    except:
        return pd.DataFrame()

# 💡 [MATH MATRIX ENGINE V20.6]
def get_math_dynamic_weightage(selected_lessons):
    base_matrix = {
        "Relations and Functions": {"1M": 1.5, "2M": 2, "5M": 2, "8M": 0},
        "Numbers and Sequences":   {"1M": 2, "2M": 2, "5M": 2, "8M": 0},
        "Algebra":                  {"1M": 2, "2M": 2, "5M": 2, "8M": 1},
        "Geometry":                 {"1M": 2, "2M": 1, "5M": 1, "8M": 1},
        "Coordinate Geometry":     {"1M": 1.5, "2M": 2, "5M": 2, "8M": 0},
        "Trigonometry":            {"1M": 1, "2M": 1, "5M": 1, "8M": 0},
        "Mensuration":             {"1M": 1.5, "2M": 2, "5M": 2, "8M": 0},
        "Statistics and Probability":{"1M": 2, "2M": 2, "5M": 2, "8M": 0}
    }
    rules = []
    for lesson in selected_lessons:
        if lesson in base_matrix:
            bm = base_matrix[lesson]
            rules.append(f"- From Chapter '{lesson}': Generate approx {int(bm['1M'])} MCQs, {bm['2M']} Questions (2-Mark), {int(bm['5M'])} Questions (5-Mark), and {bm['8M']} Question (8-Mark).")
    return "\n".join(rules)

def get_english_blueprint_rules():
    return """
    [STRICT MASTER ENGLISH BLUEPRINT LOCK]
    PART I (14 Marks): One Mark Questions (Q1-14: Synonyms, Antonyms, Grammar Grids).
    PART II (20 Marks): Two Mark Questions (Answer any 10 out of 12).
    PART III (35 Marks): Five Mark Questions (Answer any 7 out of 10).
    PART IV (16 Marks): Eight Mark Questions (Answer both - Internal Choice).
    """

# 💡 [BLUEPRINT DEFAULTS BUG FIX]: 25 மார்க்காக இருந்தாலும் பகுதி IV மதிப்புகள் 0 ஆக மாறாமல் தடுத்தல்
def get_blueprint_defaults(total_marks, is_social=False, is_english=False):
    if is_english or is_social:
        return {"p1": 14, "p2g": 12, "p2a": 10, "p3g": 10, "p3a": 7, "p4v": 8, "p4g": 2, "p4a": 2}
    defaults = {"p1": 14, "p2g": 12, "p2a": 10, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 4, "p4a": 2}
    if total_marks == 106:
        defaults = {"p1": 20, "p2g": 12, "p2a": 10, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 4, "p4a": 2}
    elif total_marks == 50:
        defaults = {"p1": 10, "p2g": 8, "p2a": 6, "p3g": 6, "p3a": 4, "p4v": 8, "p4g": 2, "p4a": 1}
    elif total_marks == 25:
        # 🛡️ பக் ஃபிக்ஸ்: கொடுக்க வேண்டியவை 2, எழுத வேண்டியவை 1 என இயல்பாக மாற்றப்பட்டுள்ளது
        defaults = {"p1": 5, "p2g": 6, "p2a": 5, "p3g": 3, "p3a": 2, "p4v": 8, "p4g": 2, "p4a": 1}
    return defaults

def generate_geometry_image(label_text=""):
    fig, ax = plt.subplots(figsize=(2.5, 2.5))
    points = np.array([[0, 0], [4, 0], [2, 3], [0, 0]])
    ax.plot(points[:, 0], points[:, 1], 'k-', lw=2)
    ax.text(-0.2, -0.2, 'A', fontsize=11, fontweight='bold')
    ax.text(4.1, -0.2, 'B', fontsize=11, fontweight='bold')
    ax.text(2, 3.2, 'C', fontsize=11, fontweight='bold')
    clean_label = label_text.replace("3/5", "Scale Factor = 3/5")
    if clean_label:
        ax.text(2, -0.6, clean_label, fontsize=9, ha='center', fontweight='bold', color='blue')
    ax.set_aspect('equal')
    ax.axis('off')
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=100)
    img_buf.seek(0)
    plt.close(fig)
    return img_buf

def generate_graph_image(label_text=""):
    fig, ax = plt.subplots(figsize=(2.5, 2.5))
    ax.axhline(0, color='black', lw=1.5)
    ax.axvline(0, color='black', lw=1.5)
    ax.grid(True, which='both', color='gray', linestyle='--', linewidth=0.5)
    ax.set_xlim([-5, 5])
    ax.set_ylim([-10, 10])
    if label_text:
        ax.set_title(label_text, fontsize=9, fontweight='bold', color='blue')
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=100)
    img_buf.seek(0)
    plt.close(fig)
    return img_buf

# ==========================================
# 3. Adaptive Language & Subject Prompt Engine
# ==========================================
def generate_prompt_v18(subject, lessons_list, exam_type, exam_time, total_marks, exam_mode, blueprint_desc, part1_val, part2_val, part3_val, diff_level):
    lessons_str = ", ".join(lessons_list)
    sub_lower = subject.lower()
    is_english = "english" in sub_lower or "ஆங்கிலம்" in sub_lower
    is_tamil = "tamil" in sub_lower or "தமிழ்" in sub_lower
    is_social = "social" in sub_lower or "சமூக" in sub_lower
    is_math = "math" in sub_lower or "கணிதம்" in sub_lower
    
    if diff_level == "எளிமை (Easy)":
        difficulty_directive = "DIFFICULTY CRITERIA: Focus 80% on direct textbook back LOTS questions."
    elif diff_level == "நடுத்தரம் (Medium)":
        difficulty_directive = "DIFFICULTY CRITERIA: Balanced public paper structure. 60% LOTS, 30% MOTS, 10% HOTS."
    else:
        difficulty_directive = "DIFFICULTY CRITERIA: High-level standard with creative HOTS problems."

    if is_math:
        subject_blueprint_rules = f"""
        [MANDATORY CRITICAL MATHEMATICS CORE EMBEDDED LOCK - VERSION 20.6]
        1. CHAPTER WEIGHTAGE DIRECTION: You MUST follow this exact chapter-wise question count allocation rule for selected chapters:
        {get_math_dynamic_weightage(lessons_list)}
        
        2. COMPULSORY QUESTION ALGORITHM:
           - Question No 28 (2-Mark Part) MUST be a Compulsory Question from Chapter 5 or Chapter 7. Creative MOTS style.
           - Question No 42 (5-Mark Part) MUST be a Compulsory Question from Chapter 2 or Chapter 3. Challenging HOTS style.
        
        3. PART IV (8-MARK) STRICT EITHER/OR PUBLIC EXAM STRUCTURE:
           - Part IV MUST contain questions with internal (either/or) choices only.
           - Graph Questions MUST include a complete X-Y Coordinate Data Table inside the text and output a separate line with the tag [DRAW_GRAPH: <Equation details>].
           - Practical Geometry Construction MUST include a separate line with the tag [DRAW_GEOMETRY: <Measurements details>]. 
           - In the Answer Key for Construction, provide the EXACT mathematically derived measurements. Do NOT use terms like 'approximately'.

        === MANDATORY SCHEME OF VALUATION FOR MATHEMATICS ANSWER KEY ===
        - 1 Mark Questions: Option Code + Exact Answer [1 Mark].
        - 2 Mark Questions: Formula/Concept [1 Mark] + Final Correct Answer [1 Mark] = [2 Marks].
        - 5 Mark Questions: Formula [1 Mark] + Substitution [1 Mark] + Working Steps [2 Marks] + Final Answer [1 Mark] = [5 Marks].
        - 8 Mark Questions: Diagram/Graph Layout [1 Mark] + Formula [1 Mark] + Substitution/Coordinates Table [2 Marks] + Calculations [2 Marks] + Final Conclusion [2 Marks] = [8 Marks].
        """
        lang_instruction = "5. Language: Pure TAMIL language only for all question text (except mathematical variables, equations, and expressions)."
        header_format = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்] (வினாக்கள் எண்ணிக்கை x மதிப்பெண் = மொத்த மதிப்பெண்கள்)"
        option_format = "Options marker: அ) , ஆ) , இ) , ஈ)"
    elif is_english:
        lang_instruction = "5. Language: Pure ENGLISH only. No Tamil markers or disclaimers in question texts."
        header_format = "PART [ROMAN_NUM] - [Section Description] (No_of_Qs x Marks = Total_Marks)"
        option_format = "Options marker: a) , b) , c) , d)"
        subject_blueprint_rules = f"""
        [STRICT TN BOARD ENGLISH BLUEPRINT LOCK]
        {get_english_blueprint_rules()}
        === MANDATORY SCHEME OF VALUATION FOR ENGLISH ANSWER KEY ===
        - One Mark Qs: Direct exact answer only [1 Mark].
        - Prose 2 Marks: Main Answer Point [1 Mark] + Supporting Point/Expansion [1 Mark] = [2 Marks].
        - Core Grammar Transformations: Direct Correct Transformation Form [2 Marks].
        - Road Map: Direction Points Guide [2 Marks].
        - Literature Paragraphs (5 Marks): Content / Relevance [3 Marks] + Organization [1 Mark] + Language [1 Mark] = [5 Marks].
        - Advertisement (5 Marks): Format/Layout [1 Mark] + Heading [1 Mark] + Content [2 Marks] + Contact Details [1 Mark] = [5 Marks].
        - Letter Writing (5 Marks): Format [1 Mark] + Content/Body [3 Marks] + Language & Closing [1 Mark] = [5 Marks].
        - Picture Description: 5 Appropriate Sentences x 1 Mark = [5 Marks].
        - Hints Development Essay (8 Marks): Introduction [1 Mark] + Development using hints [5 Marks] + Coherence & Language [2 Marks] = [8 Marks].
        """
    else:
        lang_instruction = "5. Language: Pure TAMIL language only."
        header_format = "பகுதி [ROMAN_NUM]"
        option_format = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = ""

    return f"Subject: {subject}\nLessons: {lessons_str}\nExam Type: {exam_type}\nTotal Marks: {total_marks}\nTime: {exam_time}\nMode: {exam_mode}\n{difficulty_directive}\n{blueprint_desc}\n{header_format}\n{option_format}\n{lang_instruction}\n{subject_blueprint_rules}\n=== ANSWER KEY ==="

def set_cell_margins(cell, **kwargs):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m in ["top", "start", "bottom", "end"]:
        if m in kwargs:
            node = OxmlElement(f"w:{m}")
            node.set(qn('w:w'), str(kwargs[m]))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
    tcPr.append(tcMar)

def add_solid_line(doc):
    p = doc.add_paragraph()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    p.paragraph_format.element.get_or_add_pPr().append(pBdr)

def write_markdown_to_word(doc, text):
    for line in text.split('\n'):
        line = line.strip()
        if not line: continue
        
        geo_match = re.search(r'\[DRAW_GEOMETRY[:\s]*(.*?)\]', line, re.IGNORECASE)
        if geo_match:
            label_text = geo_match.group(1)
            try:
                img_buf = generate_geometry_image(label_text)
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_buf, width=Inches(2.5))
            except Exception as e:
                doc.add_paragraph(f"[Error loading diagram: {e}]")
            continue

        graph_match = re.search(r'\[DRAW_GRAPH[:\s]*(.*?)\]', line, re.IGNORECASE)
        if graph_match:
            label_text = graph_match.group(1)
            try:
                img_buf = generate_graph_image(label_text)
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_buf, width=Inches(2.5))
            except Exception as e:
                doc.add_paragraph(f"[Error loading graph layout: {e}]")
            continue

        clean_line = line.replace('*', '').replace('$', '').strip()
        
        if "பகுதி" in clean_line or "PART" in clean_line.upper():
            marks_match = re.search(r'\(?\d+\s*[xX*]\s*\d+\s*=\s*\d+\)?', clean_line)
            if marks_match:
                calc_str = marks_match.group(0)
                title_str = clean_line.replace(calc_str, "").strip(":- ")
                table = doc.add_table(rows=1, cols=2)
                c1, c2 = table.rows[0].cells
                c1.paragraphs[0].add_run(title_str).bold = True
                c2.paragraphs[0].add_run(f"({calc_str})" if not calc_str.startswith("(") else calc_str).bold = True
                c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_cell_margins(c1, top=0, bottom=0, start=0, end=0)
                set_cell_margins(c2, top=0, bottom=0, start=0, end=0)
                continue

        option_markers = ["அ)", "ஆ)", "இ)", "ஈ)", "a)", "b)", "c)", "d)"]
        if any(marker in clean_line for marker in option_markers):
            raw_parts = re.split(r'(அ\)|ஆ\)|இ\)|ஈ\)|a\)|b\)|c\)|d\))', clean_line)
            parts = []
            current = ""
            for chunk in raw_parts:
                if chunk in option_markers:
                    if current.strip(): parts.append(current.strip())
                    current = chunk + " "
                else: current += chunk
            if current.strip(): parts.append(current.strip())
            
            if parts:
                table = doc.add_table(rows=1, cols=len(parts))
                for idx, opt in enumerate(parts):
                    cell = table.cell(0, idx)
                    cell.paragraphs[0].add_run(opt.replace("*", ""))
                    set_cell_margins(cell, top=0, bottom=0, start=0, end=0)
                continue

        p = doc.add_paragraph()
        if re.match(r'^\d+\.', clean_line):
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.first_line_indent = Inches(-0.25)
        parts = re.split(r'\*\*(.*?)\*\*', line)
        for i, part in enumerate(parts):
            run = p.add_run(part.replace('$', ''))
            if i % 2 == 1: run.bold = True

def create_professional_docx(ai_response, school_name, class_val, subject_val, exam_type, time_val, marks_val):
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.27), Inches(11.69)
    section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Inches(0.5)
    style = doc.styles['Normal']
    style.font.name = 'Nirmala UI'
    
    h_school = doc.add_paragraph(style='Normal')
    h_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_school.add_run(school_name.upper()).bold = True
    h_school.runs[0].font.size = Pt(15)
    
    table = doc.add_table(rows=2, cols=2)
    def format_cell(cell, text, align_right=False):
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if align_right else WD_ALIGN_PARAGRAPH.LEFT
        p.add_run(text).bold = True
        set_cell_margins(cell, top=0, bottom=0, start=0, end=0)
        
    format_cell(table.cell(0, 0), f"Class : {class_val}")
    format_cell(table.cell(0, 1), f"Time : {time_val}", align_right=True)
    format_cell(table.cell(1, 0), f"Subject : {subject_val}")
    format_cell(table.cell(1, 1), f"Marks : {marks_val}", align_right=True)
    
    add_solid_line(doc)
    
    if "=== ANSWER KEY ===" in ai_response:
        parts = ai_response.split("=== ANSWER KEY ===")
        write_markdown_to_word(doc, parts[0].strip())
        doc.add_page_break()
        p_key_title = doc.add_paragraph()
        p_key_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_key_title.add_run("=== ANSWER KEY & SCHEME OF VALUATION ===").bold = True
        p_key_title.runs[0].font.size = Pt(13)
        write_markdown_to_word(doc, parts[1].strip())
    else:
        write_markdown_to_word(doc, ai_response.strip())
        
    return doc

# ==========================================
# 4. Streamlit Presentation UI Config
# ==========================================
st.set_page_config(page_title="PMP AI Suite PRO", page_icon="🎓", layout="centered")

st.markdown("""
<style>
    html, body, [data-testid="stMarkdownContainer"] p {
        font-size: 19px !important;
        font-weight: 500 !important;
    }
    .stSelectbox label, .stTextInput label, .stNumberInput label {
        font-size: 20px !important;
        font-weight: bold !important;
        color: #1E3A8A !important;
    }
    button {
        font-size: 20px !important;
        height: 50px !important;
    }
</style>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🎓 வினாத்தாள் தயாரிப்பு", "📝 விடைத்தாள் திருத்தம்"])

with tab1:
    st.title("🎓 PMP Master AI Engine (V20.6)")
    df = load_data()
    if not df.empty:
        col1, col2 = st.columns(2)
        with col1:
            school_name = st.text_input("School Name", value="ABC SCHOOL")
            class_val = st.selectbox("Class", ["10"])
            subject_list = df['Subject'].unique()
            subject_val = st.selectbox("Subject", subject_list)
        with col2:
            exam_type = st.selectbox("Exam Type", ["Unit Test", "Quarterly Exam", "Half-Yearly Exam", "Annual Exam"])
            time_val = st.selectbox("Time (நேரம்)", ["1.00 Hour", "2.00 Hours", "3.00 Hours"], index=2)
            marks_val = st.number_input("Total Marks", value=100, step=1)
            
        exam_mode = st.selectbox("Exam Mode", ["🏛️ Public Exam Mode", "🏫 School Elite Mode"])
        lesson_list = df[df['Subject'] == subject_val]['Lesson'].unique()
        selected_lessons = st.multiselect("பாடங்களைத் தேர்ந்தெடுக்கவும்", lesson_list)
        
        st.markdown("---")
        diff_level = st.selectbox("வினாத்தாள் கடினத்தன்மை (Difficulty Level)", ["எளிமை (Easy)", "நடுத்தரம் (Medium)", "கடினம் (Hard)"], index=1)
        
        st.markdown("---")
        is_eng = "english" in subject_val.lower() or "ஆங்கிலம்" in subject_val.lower()
        is_soc = "social" in subject_val.lower() or "சமூக" in subject_val.lower()
        bp = get_blueprint_defaults(marks_val, is_social=is_soc, is_english=is_eng)
        
        st.markdown("### 📋 வினா வடிவமைப்பு பிரிவு")
        show_p1 = st.checkbox("பகுதி I (1-Mark Questions) சேர்க்கலாமா?", value=True)
        show_p2 = st.checkbox("பகுதி II (2-Mark Questions) சேர்க்கலாமா?", value=True)
        show_p3 = st.checkbox("பகுதி III (5-Mark Questions) சேர்க்கலாமா?", value=True)
        show_p4 = st.checkbox("பகுதி IV (Long Questions) சேர்க்கலாமா?", value=True)
        
        st.markdown("#### ⚙️ மதிப்பெண் விவரங்கள் (Fine-Tune Variables)")
        b1, b2 = st.columns(2)
        with b1:
            p1_ask = st.number_input("1-மார்க் வினாக்கள் எண்ணிக்கை", value=int(bp["p1"]) if show_p1 else 0, step=1, disabled=not show_p1)
            p2_get = st.number_input("2-மார்க் கொடுக்க வேண்டியவை (Given)", value=int(bp["p2g"]) if show_p2 else 0, step=1, disabled=not show_p2)
            p2_ask = st.number_input("2-மார்க் எழுத வேண்டியவை (Answer)", value=int(bp["p2a"]) if show_p2 else 0, step=1, disabled=not show_p2)
        with b2:
            p3_get = st.number_input("5-மார்க் கொடுக்க வேண்டியவை (Given)", value=int(bp["p3g"]) if show_p3 else 0, step=1, disabled=not show_p3)
            p3_ask = st.number_input("5-மார்க் எழுத வேண்டியவை (Answer)", value=int(bp["p3a"]) if show_p3 else 0, step=1, disabled=not show_p3)
            p4_val = st.selectbox("நெடுவினா மதிப்பெண்", [5, 8, 10], index=1 if is_eng or is_soc or marks_val==100 else 0, disabled=not show_p4)
            p4_get = st.number_input("நெடுவினா கொடுக்க வேண்டியவை (Given)", value=int(bp["p4g"]) if show_p4 else 0, step=1, disabled=not show_p4)
            p4_ask = st.number_input("நெடுவினா எழுத வேண்டியவை (Answer)", value=int(bp["p4a"]) if show_p4 else 0, step=1, disabled=not show_p4)
            
        total_calculated = (p1_ask * 1) + (p2_ask * 2) + (p3_ask * 5) + (p4_ask * p4_val)
        can_generate = total_calculated == marks_val
        
        if can_generate:
            st.success(f"✅ மதிப்பெண்கள் சரியாகப் பொருந்தியது: {total_calculated} மார்க்.")
        else:
            st.warning(f"⚠️ கணக்கீடு: {total_calculated} மார்க் | மொத்த மதிப்பெண்: {marks_val} மார்க். (தயவுசெய்து சமப்படுத்தவும்).")
            
        if st.button("🚀 Generate PRO Question Paper", use_container_width=True):
            if can_generate and selected_lessons:
                timer_placeholder = st.empty()
                start_time = time.time()
                
                with st.spinner("⏳ அசல் அரசுப் பொதுத்தேர்வு விதிகளின்படி வினாத்தாள் தயாராகிறது..."):
                    blueprint_desc = f"- Part I: {p1_ask} Qs. - Part II: Given {p2_get}, Answer {p2_ask}. - Part III: Given {p3_get}, Answer {p3_ask}. - Part IV: Given {p4_get}, Answer {p4_ask}."
                    prompt = generate_prompt_v18(subject_val, selected_lessons, exam_type, time_val, marks_val, exam_mode, blueprint_desc, p1_ask, p2_ask, p3_ask, diff_level)
                    
                    response = None
                    max_retries = 4
                    for attempt in range(max_retries):
                        try:
                            response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                            break
                        except Exception as api_err:
                            if ("429" in str(api_err) or "503" in str(api_err)) and attempt < max_retries - 1:
                                wait_time = (attempt + 1) * 4
                                time.sleep(wait_time)
                                continue
                            else:
                                st.error(f"சர்வர் தற்காலிகமாக ஓவர்லோடு ஆகியுள்ளது: {api_err}")
                    
                    elapsed_time = time.time() - start_time
                    timer_placeholder.info(f"⏱️ வினாத்தாள் உருவாக்க எடுக்கப்பட்ட நேரம்: {elapsed_time:.1f} விநாடிகள் (Seconds)")
                                
                    if response:
                        doc = create_professional_docx(response.text, school_name, class_val, subject_val, exam_type, time_val, marks_val)
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        st.session_state['docx_bytes'] = doc_io.getvalue()
                        st.session_state['preview_text'] = response.text
                        st.success("✅ வினாத்தாள் வெற்றிகரமாகத் தயாராகிவிட்டது!")

        if 'preview_text' in st.session_state:
            st.markdown("---")
            st.subheader("👀 வினாத்தாள் மற்றும் விடைக்குறிப்பு முன்னோட்டம் (Live Preview)")
            with st.container(border=True):
                st.markdown(st.session_state['preview_text'])
            st.markdown("---")

        if 'docx_bytes' in st.session_state:
            st.download_button(label="📥 Download Word File (.docx)", data=st.session_state['docx_bytes'], file_name=f"PMP_{subject_val}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

with tab2:
    st.title("📝 AI Math Paper Evaluator (Multi-Format Edition)")
    uploaded_file = st.file_uploader("உங்கள் விடைத்தாளைத் தேர்ந்தெடுக்கவும் (Image / PDF)", type=["png", "jpg", "jpeg", "pdf"])
    
    if uploaded_file is not None:
        file_type = uploaded_file.name.split(".")[-1].lower()
        eval_payload = []
        
        if file_type == "pdf":
            st.info("📊 PDF கோப்பு கண்டறியப்பட்டது. ஜெமினி ஏஐ பகுப்பாய்வு செய்கிறது...")
            file_data = uploaded_file.read()
            pdf_part = types.Part.from_bytes(data=file_data, mime_type="application/pdf")
            eval_payload.append(pdf_part)
        else:
            image = Image.open(uploaded_file)
            st.image(image, caption="Uploaded Handwriting Page", width=420)
            eval_payload.append(image)
            
        if st.button("🚀 Start AI Evaluation", use_container_width=True):
            with st.spinner("⏳ ஜெமினி AI விடைத்தாளைத் திருத்தி வருகிறது..."):
                eval_prompt = "You are an official TN Board Math Evaluator. Read handwriting or PDF pages and correct step-by-step. Write fractions as $\\frac{a}{b}$ inside single dollar signs. Respond in Tamil."
                
                eval_payload.append(eval_prompt)
                response = None
                for attempt in range(3):
                    try:
                        response = client.models.generate_content(model='gemini-2.5-flash', contents=eval_payload)
                        break
                    except Exception as eval_err:
                        if "429" in str(eval_err) and attempt < 2:
                            time.sleep(4)
                            continue
                        else:
                            st.error(f"மதிப்பீட்டு சர்வர் பிழை: {eval_err}")
                            
                if response:
                    st.markdown(response.text)
