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

# 💡 அசல் கணித அலகுகள் வாரியான வினாக்கள் விபரத் தரவுத்தளம் (Frequency Matrix Engine)
def get_math_dynamic_weightage(selected_lessons, part1_val, part2_val, part3_val):
    base_matrix = {
        "Relations and Functions": {"1M": 1.5, "2M": 2, "5M": 1.5, "8M": 0},
        "Numbers and Sequences":   {"1M": 2.0, "2M": 2, "5M": 2.0, "8M": 0},
        "Algebra":                  {"1M": 2.0, "2M": 2, "5M": 2.0, "8M": 1},
        "Geometry":                 {"1M": 2.0, "2M": 1, "5M": 1.0, "8M": 1},
        "Coordinate Geometry":     {"1M": 1.5, "2M": 2, "5M": 2.0, "8M": 0},
        "Trigonometry":            {"1M": 1.0, "2M": 1, "5M": 1.0, "8M": 0},
        "Mensuration":             {"1M": 1.5, "2M": 2, "5M": 2.0, "8M": 0},
        "Statistics and Probability":{"1M": 2.0, "2M": 2, "5M": 2.0, "8M": 0}
    }
    rules = []
    for lesson in selected_lessons:
        if lesson in base_matrix:
            bm = base_matrix[lesson]
            rules.append(f"- From '{lesson}': Generate approx {int(bm['1M'])} MCQs, {bm['2M']} Questions (2-Mark), {int(bm['5M'])} Questions (5-Mark), and {bm['8M']} Question (8-Mark).")
    return "\n".join(rules)

# 💡 மாஸ்டர் ஆங்கில போர்டு ப்ளூபிரின்ட் இன்ஜின் (V3 Base - 100% Balanced Literature Filter)
def get_english_blueprint_rules():
    return """
    [STRICT MASTER ENGLISH BLUEPRINT LOCK]
    PART I (14 Marks): One Mark Questions
    - Q1-3: Synonyms strictly selected from the text prose context (e.g., plaintively, dilated, yanked).
    - Q4-6: Antonyms strictly selected from the text prose context (e.g., cranky, commotion).
    - Q7-14: Textual Grammar Matrix: Plural Form, Suffix/Prefix, Abbreviations, Phrasal Verbs, Compound Words, Prepositions, Tense Forms, Linkers.

    PART II (20 Marks): Two Mark Questions (Answer any 10 out of 12)
    - Section 1 (Prose Qs): 3 Short answer questions from selected textbook lessons.
    - Section 2 (Poetry Appreciation): 3 Poetic line sets with internal appreciation questions.
    - Section 3 (Grammar Blocks): 5 Mandatory Core Grammar Questions: Voice Change, Reported Speech, Punctuation, Sentence Transformation, Rearrange jumbled words.
    - Section 4 (Roadmap): 1 Question on Road Map directions guide.

    PART III (35 Marks): Five Mark Questions (Answer any 7 out of 10)
    - Section 1 (Literature Paragraphs): Prose Paragraphs and Poetry Paragraphs covering textbook lessons.
    - Section 2 (Coherent Order & Comprehension): Rearrange sentences in coherent order, Supplementary passage comprehension.
    - Section 3 (Writing Skills Matrix): Advertisement, Formal Letter, Notice Writing, Picture Comprehension.

    PART IV (16 Marks): Eight Mark Questions (Answer both - Internal Choice)
    - Q37: Comprehensive Prose Passage reading or Poem reading with granularity questions.
    - Q38: Detailed Literature Essay from supplementary stories.
    """

def get_blueprint_defaults(total_marks, is_social=False, is_english=False):
    if is_english or is_social:
        return {"p1": 14, "p2g": 12, "p2a": 10, "p3g": 10, "p3a": 7, "p4v": 8, "p4g": 2, "p4a": 2}
    defaults = {"p1": 14, "p2g": 12, "p2a": 10, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 4, "p4a": 2}
    if total_marks == 106:
        defaults = {"p1": 20, "p2g": 12, "p2a": 10, "p3g": 14, "p3a": 10, "p4v": 8, "p4g": 4, "p4a": 2}
    elif total_marks == 50:
        defaults = {"p1": 10, "p2g": 8, "p2a": 6, "p3g": 6, "p3a": 4, "p4v": 8, "p4g": 2, "p4a": 1}
    elif total_marks == 25:
        defaults = {"p1": 5, "p2g": 6, "p2a": 5, "p3g": 3, "p3a": 2, "p4v": 8, "p4g": 0, "p4a": 0}
    return defaults

def generate_geometry_image(shape_type, label_text=""):
    fig, ax = plt.subplots(figsize=(2.5, 2.5))
    shape_upper = shape_type.upper()
    clean_label = label_text.replace("Angle", r"$\angle$").replace("angle", r"$\angle$")
    if "TRIANGLE" in shape_upper:
        points = np.array([[0, 0], [4, 0], [2, 3], [0, 0]])
        ax.plot(points[:, 0], points[:, 1], 'k-', lw=2)
        ax.text(-0.2, -0.2, 'A', fontsize=11, fontweight='bold')
        ax.text(4.1, -0.2, 'B', fontsize=11, fontweight='bold')
        ax.text(2, 3.2, 'C', fontsize=11, fontweight='bold')
        if clean_label:
            ax.text(2, -0.6, clean_label, fontsize=10, ha='center', fontweight='bold', color='blue')
    elif "SQUARE" in shape_upper:
        points = np.array([[0, 0], [3, 0], [3, 3], [0, 3], [0, 0]])
        ax.plot(points[:, 0], points[:, 1], 'k-', lw=2)
        ax.text(-0.2, -0.2, 'A', fontsize=10)
        ax.text(3.2, -0.2, 'B', fontsize=10)
        if clean_label:
            ax.text(1.5, -0.6, clean_label, fontsize=10, ha='center', fontweight='bold', color='blue')
    ax.set_aspect('equal')
    ax.axis('off')
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
        difficulty_directive = "DIFFICULTY CRITERIA: Focus 80% on direct textbook back questions and formulas."
    elif diff_level == "நடுத்தரம் (Medium)":
        difficulty_directive = "DIFFICULTY CRITERIA: Balanced public paper structure. 60% Direct, 30% Application, 10% HOTS."
    else:
        difficulty_directive = "DIFFICULTY CRITERIA: High-level revision standard with indirect grammatical transformations."

    math_weightage_directive = ""
    if is_math:
        math_weightage_directive = f"[STRICT LESSON-WISE MARKS WEIGHTAGE MATRIX]\n{get_math_dynamic_weightage(lessons_list, part1_val, part2_val, part3_val)}"

    if is_english:
        lang_instruction = "5. Language: Pure ENGLISH only. No Tamil markers or disclaimers in question texts."
        header_format = "PART [ROMAN_NUM] - [Section Description] (No_of_Qs x Marks = Total_Marks)"
        option_format = "Options marker: a) , b) , c) , d)"
        subject_blueprint_rules = f"[STRICT TN BOARD ENGLISH BLUEPRINT LOCK]\n{get_english_blueprint_rules()}"
    elif is_tamil:
        lang_instruction = "5. Language: Pure TAMIL only."
        header_format = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்] (வினாக்கள் எண்ணிக்கை x மதிப்பெண் = மொத்த மதிப்பெண்கள்)"
        option_format = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = "[அசல் தமிழ் பாடத்திட்ட ப்ளூபிரின்ட்] சொல்வளம், இலக்கணம் லாக்."
    elif is_social:
        lang_instruction = "5. Language: Pure TAMIL only."
        header_format = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்]"
        option_format = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = "[MANDATORY CRITICAL SOCIAL SCIENCE BLUEPRINT] Assertion-Reason, Map locked."
    elif is_math:
        lang_instruction = "5. Language: Pure TAMIL only."
        header_format = "பகுதி [ROMAN_NUM] - [பிரிவின் விளக்கம்] (No_of_Qs x Marks = Total_Marks)"
        option_format = "Options marker: அ) , ஆ) , இ) , ஈ)"
        subject_blueprint_rules = f"[MANDATORY CRITICAL MATHEMATICS CORE EMBEDDED LOCK]\n1. ABSOLUTE BAN ON AI DISCLAIMERS.\n2. DYNAMIC GEOMETRY TAGS: [DRAW_TRIANGLE: AB=5cm].\n3. GRAPH PAPER COORDINATES.\n{math_weightage_directive}"
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
        
        draw_match = re.search(r'\[DRAW_(TRIANGLE|SQUARE)[:\s]*(.*?)\]', line, re.IGNORECASE)
        if draw_match:
            shape_type = draw_match.group(1)
            label_text = draw_match.group(2)
            try:
                img_buf = generate_geometry_image(shape_type, label_text)
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_buf, width=Inches(2.5))
            except Exception as e:
                doc.add_paragraph(f"[Error loading diagram: {e}]")
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
    parts = ai_response.split("=== ANSWER KEY ===")
    write_markdown_to_word(doc, parts[0].strip())
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
    st.title("🎓 PMP Master AI Engine (V18.8)")
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
        
        st.markdown("### 📋 வினா வடிவமைப்பு பிரிவு (பகுதிகளைத் தேர்ந்தெடுக்கவும்)")
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
                with st.spinner("⏳ வினாத்தாள் தயாராகிறது..."):
                    blueprint_desc = f"- Part I: {p1_ask} Qs. - Part II: Given {p2_get}, Answer {p2_ask}. - Part III: Given {p3_get}, Answer {p3_ask}. - Part IV: Given {p4_get}, Answer {p4_ask}."
                    prompt = generate_prompt_v18(subject_val, selected_lessons, exam_type, time_val, marks_val, exam_mode, blueprint_desc, p1_ask, p2_ask, p3_ask, diff_level)
                    
                    response = None
                    for attempt in range(3):
                        try:
                            response = client.models.generate_content(model='gemini-2.5-pro', contents=prompt)
                            break
                        except:
                            time.sleep(2)
                            continue
                            
                    if response:
                        doc = create_professional_docx(response.text, school_name, class_val, subject_val, exam_type, time_val, marks_val)
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        st.session_state['docx_bytes'] = doc_io.getvalue()
                        st.success("✅ வினாத்தாள் வெற்றிகரமாகத் தயாராகிவிட்டது!")

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
                try:
                    eval_payload.append(eval_prompt)
                    response = client.models.generate_content(model='gemini-2.5-pro', contents=eval_payload)
                    st.markdown(response.text)
                except Exception as e:
                    st.error(f"❌ சர்வர் பிழை: {e}")
